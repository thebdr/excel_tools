using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelPowerTools.Interop;

namespace ExcelPowerTools.UI
{
    internal enum CsvConflictAction
    {
        Cancel,
        Reload,
        SaveCopy,
        Overwrite,
    }

    /// <summary>Explicit choices for a source that changed or disappeared after it was opened.</summary>
    internal sealed class CsvConflictForm : Form
    {
        public CsvConflictAction SelectedAction { get; private set; } = CsvConflictAction.Cancel;

        public static CsvConflictAction Choose(
            string sheetName,
            string path,
            DateTime openedModifiedLocal,
            DateTime? currentModifiedLocal,
            bool sourceExists,
            bool hasLocalChanges,
            bool hasViewOnlyContent)
        {
            using (var form = new CsvConflictForm(
                sheetName, path, openedModifiedLocal, currentModifiedLocal,
                sourceExists, hasLocalChanges, hasViewOnlyContent))
            {
                form.ShowDialog(ExcelWindow.Current);
                return form.SelectedAction;
            }
        }

        private CsvConflictForm(
            string sheetName,
            string path,
            DateTime openedModifiedLocal,
            DateTime? currentModifiedLocal,
            bool sourceExists,
            bool hasLocalChanges,
            bool hasViewOnlyContent)
        {
            Text = sourceExists ? "File changed outside Excel" : "Source file is missing";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(680, 300);
            Font = new Font("Segoe UI", 9f);

            var heading = new Label
            {
                Text = sourceExists
                    ? "The source changed after this sheet was opened."
                    : "The source path no longer exists.",
                Font = new Font(Font, FontStyle.Bold),
                Location = new Point(18, 16),
                Size = new Size(640, 24),
            };
            Controls.Add(heading);

            var details = new TextBox
            {
                ReadOnly = true,
                Multiline = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = SystemColors.Window,
                Location = new Point(18, 48),
                Size = new Size(644, 132),
                Text =
                    "Sheet: " + sheetName + Environment.NewLine +
                    "Path: " + path + Environment.NewLine +
                    "Modified when opened: " + FormatDate(openedModifiedLocal) + Environment.NewLine +
                    "Modified now: " + (sourceExists ? FormatDate(currentModifiedLocal) : "missing") + Environment.NewLine +
                    "Local source-value changes: " + (hasLocalChanges ? "yes" : "no") + Environment.NewLine +
                    "View-only helper content: " + (hasViewOnlyContent ? "yes (Reload will discard it)" : "no"),
            };
            Controls.Add(details);

            var warning = new Label
            {
                Text = sourceExists
                    ? "Reload rebuilds the sheet from disk. Overwrite replaces the disk file and preserves it as .bak."
                    : "Save a copy is safest. Recreate writes a new file at the missing source path.",
                Location = new Point(18, 190),
                Size = new Size(644, 36),
            };
            Controls.Add(warning);

            int y = 244;
            var reload = AddActionButton("Reload from disk", 18, y, 138, CsvConflictAction.Reload);
            reload.Enabled = sourceExists;
            AddActionButton("Save a copy...", 164, y, 126, CsvConflictAction.SaveCopy);
            AddActionButton(sourceExists ? "Overwrite anyway" : "Recreate file", 298, y, 140, CsvConflictAction.Overwrite);
            var cancel = AddActionButton("Cancel", 552, y, 110, CsvConflictAction.Cancel);
            CancelButton = cancel;
        }

        private Button AddActionButton(string text, int x, int y, int width, CsvConflictAction action)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 32),
            };
            button.Click += (_, __) =>
            {
                SelectedAction = action;
                DialogResult = action == CsvConflictAction.Cancel ? DialogResult.Cancel : DialogResult.OK;
                Close();
            };
            Controls.Add(button);
            return button;
        }

        private static string FormatDate(DateTime? value) =>
            value.HasValue && value.Value != DateTime.MinValue
                ? value.Value.ToString("yyyy-MM-dd HH:mm:ss")
                : "unknown";
    }
}
