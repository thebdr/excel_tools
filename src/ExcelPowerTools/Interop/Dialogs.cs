using System.Windows.Forms;

namespace ExcelPowerTools.Interop
{
    /// <summary>Message boxes owned by the Excel window (the add-in replacement for VBA's MsgBox).</summary>
    internal static class Dialogs
    {
        public const string Title = "Power Tools";

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

        public static string SaveFile(string filter, string defaultName)
        {
            using (var d = new SaveFileDialog { Filter = filter, FileName = defaultName, OverwritePrompt = true })
                return d.ShowDialog(ExcelWindow.Current) == DialogResult.OK ? d.FileName : null;
        }
    }
}
