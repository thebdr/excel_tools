using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using ExcelPowerTools.Settings;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// WinForms replacement for the VBA "ufIoListAuxParameters" userform: collects the separators,
    /// the field index, the new value (with the +/- add-subtract convention) and optional colours.
    /// </summary>
    internal sealed class FieldEditorForm : Form
    {
        private readonly TextBox _sep1 = new TextBox();
        private readonly TextBox _sep2 = new TextBox();
        private readonly NumericUpDown _fieldIndex = new NumericUpDown();
        private readonly TextBox _newValue = new TextBox();
        private readonly TextBox _fontColor = new TextBox();
        private readonly TextBox _interiorColor = new TextBox();

        // Six quick-pick colours (RRGGBB) offered next to the colour boxes.
        private static readonly string[] Palette = { "FF0000", "FFC000", "FFFF00", "00B050", "0070C0", "000000" };

        // Shared eraser icon for the "clear" buttons, and a tooltip so the icon-only buttons stay discoverable.
        private static readonly Image EraserImage = CreateEraserIcon(16);
        private readonly ToolTip _toolTip = new ToolTip();

        // Separators are returned with escape tokens translated to real characters (\s space, \n LF, \r CR, \t tab).
        public string Separator => Unescape(_sep1.Text);
        public string Separator2 => Unescape(_sep2.Text);
        public int FieldIndex => (int)_fieldIndex.Value;   // 1-based
        public string NewValue => _newValue.Text;
        public string FontColor => _fontColor.Text.Trim();
        public string InteriorColor => _interiorColor.Text.Trim();

        public FieldEditorForm(FieldEditorSettings settings)
        {
            Text = "Change Field — parameters";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            KeyPreview = true;
            ClientSize = new Size(600, 284);
            Font = new Font("Courier New", 9f);

            const int labelX = 12, inputX = 190, sepBoxW = 46, btnX = 246, colorBoxW = 110;
            int swatchX = inputX + colorBoxW + 12;
            int y = 16;

            AddLabel("Separator (start)", labelX, y);
            Place(_sep1, inputX, y, sepBoxW);
            AddSeparatorButtons(btnX, y, _sep1);
            y += 34;

            AddLabel("Separator (end, opt.)", labelX, y);
            Place(_sep2, inputX, y, sepBoxW);
            AddSeparatorButtons(btnX, y, _sep2);
            y += 34;

            AddLabel("Field index (from 1)", labelX, y);
            _fieldIndex.Minimum = 1;
            _fieldIndex.Maximum = 64;
            Place(_fieldIndex, inputX, y, 60);
            AddFieldIndexButtons(inputX + 66, y);   // quick-pick 1..8
            y += 34;

            AddLabel("New value\n(+n / -n / text)", labelX, y, 34);
            Place(_newValue, inputX, y, 130);
            AddDeltaButtons(inputX + 138, y);       // -1 / +1 step the add/subtract amount
            _newValue.KeyDown += (_, e) =>          // Up/Down arrows step the amount too
            {
                if (e.KeyCode == Keys.Up) { BumpNewValue(+1); e.Handled = e.SuppressKeyPress = true; }
                else if (e.KeyCode == Keys.Down) { BumpNewValue(-1); e.Handled = e.SuppressKeyPress = true; }
            };
            y += 44;

            AddLabel("Font colour (hex/auto)", labelX, y);
            Place(_fontColor, inputX, y, colorBoxW);
            AddColorSwatches(swatchX, y, _fontColor);
            y += 32;

            AddLabel("Fill colour (hex/none)", labelX, y);
            Place(_interiorColor, inputX, y, colorBoxW);
            AddColorSwatches(swatchX, y, _interiorColor);
            y += 34;

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Size = new Size(90, 30), Location = new Point(ClientSize.Width - 200, y) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 30), Location = new Point(ClientSize.Width - 102, y) };
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;

            // load persisted values (raw text, so escape tokens survive a round-trip)
            _sep1.Text = settings.Separator;
            _sep2.Text = settings.Separator2;
            _fieldIndex.Value = Clamp(settings.FieldIndex, (int)_fieldIndex.Minimum, (int)_fieldIndex.Maximum);
            _fontColor.Text = settings.FontColor;
            _interiorColor.Text = settings.InteriorColor;

            Shown += (_, __) => { _newValue.Focus(); };
        }

        /// <summary>Persists the current dialog values back to <paramref name="settings"/> (called on OK).</summary>
        public void CopyInto(FieldEditorSettings settings)
        {
            settings.Separator = _sep1.Text;       // raw, keeps \n / \r etc. as editable tokens
            settings.Separator2 = _sep2.Text;
            settings.FieldIndex = FieldIndex;
            settings.FontColor = FontColor;
            settings.InteriorColor = InteriorColor;
        }

        // A row of quick-insert buttons (to the RIGHT of the separator box) that SET the box.
        // \s / \n / \r / \r\n insert the readable token; Unescape() turns them into real characters.
        private void AddSeparatorButtons(int x, int y, TextBox target)
        {
            var items = new (string label, string value, int width)[]
            {
                ("(", "(", 22), (")", ")", 22), (".", ".", 22), (",", ",", 22),
                ("-", "-", 22), ("_", "_", 22), (";", ";", 22),
                ("\\s", "\\s", 28), ("\\n", "\\n", 28), ("\\r", "\\r", 28), ("\\r\\n", "\\r\\n", 48),
            };

            int bx = x;
            foreach (var it in items)
            {
                string val = it.value;
                var b = new Button
                {
                    Text = it.label,
                    Location = new Point(bx, y - 1),
                    Size = new Size(it.width, 26),
                    TabStop = false,
                    Margin = new Padding(0),
                };
                b.Click += (_, __) => target.Text = val;
                Controls.Add(b);
                bx += it.width + 2;
            }

            // Clear button: eraser icon instead of text, with a "Clear" tooltip.
            var clear = new Button
            {
                Location = new Point(bx, y - 1),
                Size = new Size(30, 26),
                TabStop = false,
                Image = EraserImage,
                ImageAlign = ContentAlignment.MiddleCenter,
            };
            clear.Click += (_, __) => target.Text = "";
            _toolTip.SetToolTip(clear, "Clear");
            Controls.Add(clear);
        }

        // Quick-pick buttons 1..8 (to the RIGHT of the field-index box) that set the field index.
        private void AddFieldIndexButtons(int x, int y)
        {
            int bx = x;
            for (int n = 1; n <= 8; n++)
            {
                int captured = n;
                var b = new Button
                {
                    Text = n.ToString(CultureInfo.InvariantCulture),
                    Location = new Point(bx, y - 1),
                    Size = new Size(24, 26),
                    TabStop = false,
                };
                b.Click += (_, __) => _fieldIndex.Value = captured;
                Controls.Add(b);
                bx += 26;
            }
        }

        // -1 / +1 buttons that step the add/subtract amount in the New value box (e.g. +1 twice -> "+2").
        private void AddDeltaButtons(int x, int y)
        {
            var minus = new Button { Text = "-1", Location = new Point(x, y - 1), Size = new Size(36, 26), TabStop = false };
            minus.Click += (_, __) => BumpNewValue(-1);

            var plus = new Button { Text = "+1", Location = new Point(x + 40, y - 1), Size = new Size(36, 26), TabStop = false };
            plus.Click += (_, __) => BumpNewValue(+1);

            Controls.Add(minus);
            Controls.Add(plus);
        }

        private void BumpNewValue(int step)
        {
            string t = _newValue.Text.Trim();
            if (t.StartsWith("+")) t = t.Substring(1);
            int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out int current);   // non-numeric -> 0
            int next = current + step;
            if (next > 500) next = 500;
            if (next < -500) next = -500;
            _newValue.Text = next > 0 ? "+" + next.ToString(CultureInfo.InvariantCulture)
                                      : next.ToString(CultureInfo.InvariantCulture);
        }

        // Six colour swatches to the RIGHT of a colour box; clicking one fills the box with its RRGGBB hex.
        private void AddColorSwatches(int x, int y, TextBox target)
        {
            int bx = x;
            foreach (string hex in Palette)
            {
                string h = hex;
                var b = new Button
                {
                    Location = new Point(bx, y - 1),
                    Size = new Size(24, 24),
                    TabStop = false,
                    FlatStyle = FlatStyle.Flat,
                    BackColor = ColorFromHex(h),
                };
                b.FlatAppearance.BorderSize = 1;
                b.Click += (_, __) => target.Text = h;
                Controls.Add(b);
                bx += 26;
            }
        }

        /// <summary>Translates the escape tokens the buttons insert (\s \n \r \t \\) into real characters.</summary>
        private static string Unescape(string s)
        {
            if (string.IsNullOrEmpty(s) || s.IndexOf('\\') < 0) return s;

            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '\\' && i + 1 < s.Length)
                {
                    char n = s[i + 1];
                    switch (n)
                    {
                        case 's': sb.Append(' '); i++; continue;
                        case 'n': sb.Append('\n'); i++; continue;
                        case 'r': sb.Append('\r'); i++; continue;
                        case 't': sb.Append('\t'); i++; continue;
                        case '\\': sb.Append('\\'); i++; continue;
                    }
                }
                sb.Append(s[i]);
            }
            return sb.ToString();
        }

        /// <summary>Draws a small angled pencil-eraser icon (pink band + light body) so it renders without an icon font.</summary>
        private static Image CreateEraserIcon(int size)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TranslateTransform(size / 2f, size / 2f);
                g.RotateTransform(-30f);

                float w = size * 0.9f, h = size * 0.55f;
                float x = -w / 2f, top = -h / 2f, bandH = h * 0.42f;

                using (var body = new SolidBrush(Color.FromArgb(238, 238, 238)))
                using (var band = new SolidBrush(Color.FromArgb(224, 64, 96)))
                using (var pen = new Pen(Color.FromArgb(80, 80, 80), 1f))
                {
                    g.FillRectangle(body, x, top + bandH, w, h - bandH);
                    g.FillRectangle(band, x, top, w, bandH);
                    g.DrawRectangle(pen, x, top, w, h);
                }
            }
            return bmp;
        }

        private static Color ColorFromHex(string hex)
        {
            return Color.FromArgb(
                Convert.ToInt32(hex.Substring(0, 2), 16),
                Convert.ToInt32(hex.Substring(2, 2), 16),
                Convert.ToInt32(hex.Substring(4, 2), 16));
        }

        private void AddLabel(string text, int x, int y, int height = 20)
        {
            Controls.Add(new Label { Text = text, Location = new Point(x, y + 3), Size = new Size(178, height), AutoSize = false });
        }

        private void Place(Control c, int x, int y, int width)
        {
            c.Location = new Point(x, y);
            c.Size = new Size(width, 24);
            Controls.Add(c);
        }

        private static int Clamp(int v, int lo, int hi) => v < lo ? lo : (v > hi ? hi : v);
    }
}
