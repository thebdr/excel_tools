using System.Drawing;
using System.Windows.Forms;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Prompts for a CSV separator with quick-insert buttons. The chosen token is returned via
    /// <see cref="RawToken"/> (caller persists it for next time); \t = tab, \s = space.
    /// </summary>
    internal sealed class CsvSeparatorForm : Form
    {
        private readonly TextBox _box = new TextBox();

        public string RawToken => _box.Text;

        public CsvSeparatorForm(string title, string initialToken)
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            KeyPreview = true;
            ClientSize = new Size(322, 132);
            Font = new Font("Segoe UI", 9f);

            Controls.Add(new Label { Text = "Separator:", Location = new Point(12, 16), AutoSize = true });
            _box.Location = new Point(92, 12);
            _box.Size = new Size(60, 24);
            _box.Text = string.IsNullOrEmpty(initialToken) ? ";" : initialToken;
            Controls.Add(_box);

            int bx = 12;
            foreach (var item in new[] { (";", ";"), (",", ","), ("|", "|"), ("Tab", "\\t"), ("Space", "\\s") })
            {
                string token = item.Item2;
                var b = new Button { Text = item.Item1, Location = new Point(bx, 50), Size = new Size(56, 26), TabStop = false };
                b.Click += (_, __) => _box.Text = token;
                Controls.Add(b);
                bx += 58;
            }

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new Point(142, 92), Size = new Size(80, 28) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(230, 92), Size = new Size(80, 28) };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        /// <summary>Translates a token (\t, \s, or a literal char) into the separator character.</summary>
        public static char Translate(string token)
        {
            if (string.IsNullOrEmpty(token)) return ';';
            if (token == "\\t") return '\t';
            if (token == "\\s") return ' ';
            return token[0];
        }
    }
}
