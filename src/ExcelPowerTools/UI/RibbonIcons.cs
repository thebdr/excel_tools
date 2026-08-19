using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace ExcelPowerTools.UI
{
    /// <summary>
    /// Small, high-contrast brand icons for the ribbon's primary actions. Secondary actions keep their
    /// familiar Office icons; these marks establish the visual hierarchy the old ribbon was missing.
    /// </summary>
    internal static class RibbonIcons
    {
        private static readonly Dictionary<string, Bitmap> Cache =
            new Dictionary<string, Bitmap>(StringComparer.OrdinalIgnoreCase);

        public static Bitmap Get(string imageId)
        {
            if (string.IsNullOrWhiteSpace(imageId)) return null;

            lock (Cache)
            {
                if (Cache.TryGetValue(imageId, out Bitmap cached)) return cached;
                Bitmap icon = Create(imageId, 32);
                if (icon != null) Cache[imageId] = icon;
                return icon;
            }
        }

        public static Bitmap Create(string imageId, int size)
        {
            switch (imageId)
            {
                case "ptHubIcon": return DrawHub(size);
                case "ptCompareIcon": return DrawCompare(size);
                case "ptOpenIcon": return DrawOpen(size);
                case "ptSaveIcon": return DrawSave(size);
                case "ptValuesIcon": return DrawValues(size);
                case "ptSaveAllIcon": return DrawSaveAll(size);
                case "ptSaveAllIcon16": return DrawSaveAll(16);
                case "ptIndexIcon": return DrawIndex(size);
                case "ptIndexIcon16": return DrawIndex(16);
                case "ptTimelineIcon": return DrawTimeline(size);
                case "ptTimelineIcon16": return DrawTimeline(16);
                default: return null;
            }
        }

        private static Bitmap Canvas(int size, Color start, Color end, Action<Graphics, float> draw)
        {
            var bitmap = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                float u = size / 32f;
                using (GraphicsPath path = RoundedRect(new RectangleF(u, u, size - 2f * u, size - 2f * u), 7f * u))
                using (var brush = new LinearGradientBrush(
                           new PointF(0, 0), new PointF(size, size), start, end))
                    g.FillPath(brush, path);
                draw(g, u);
            }
            return bitmap;
        }

        private static Bitmap DrawHub(int size)
        {
            return Canvas(size, Color.FromArgb(59, 70, 229), Color.FromArgb(111, 67, 209), (g, u) =>
            {
                Color soft = Color.FromArgb(225, 235, 255);
                using (var white = new SolidBrush(Color.White))
                using (var pale = new SolidBrush(soft))
                {
                    FillRound(g, white, 7 * u, 7 * u, 7 * u, 7 * u, 2 * u);
                    FillRound(g, pale, 18 * u, 7 * u, 7 * u, 7 * u, 2 * u);
                    FillRound(g, pale, 7 * u, 18 * u, 7 * u, 7 * u, 2 * u);
                    FillRound(g, white, 18 * u, 18 * u, 7 * u, 7 * u, 2 * u);
                }
                using (var pen = new Pen(Color.FromArgb(115, 255, 255, 255), 1.3f * u))
                {
                    g.DrawLine(pen, 21.5f * u, 17 * u, 21.5f * u, 26 * u);
                    g.DrawLine(pen, 17 * u, 21.5f * u, 26 * u, 21.5f * u);
                }
            });
        }

        private static Bitmap DrawCompare(int size)
        {
            return Canvas(size, Color.FromArgb(102, 70, 203), Color.FromArgb(218, 71, 132), (g, u) =>
            {
                using (var back = new SolidBrush(Color.FromArgb(185, 223, 255)))
                    FillRound(g, back, 5.5f * u, 7 * u, 13 * u, 18 * u, 2.5f * u);
                using (var front = new SolidBrush(Color.White))
                    FillRound(g, front, 13.5f * u, 5 * u, 13 * u, 18 * u, 2.5f * u);
                using (var grid = new Pen(Color.FromArgb(112, 91, 201), 1.2f * u))
                {
                    g.DrawLine(grid, 16.5f * u, 10 * u, 23.5f * u, 10 * u);
                    g.DrawLine(grid, 16.5f * u, 14 * u, 23.5f * u, 14 * u);
                    g.DrawLine(grid, 16.5f * u, 18 * u, 21 * u, 18 * u);
                }
                using (var badge = new SolidBrush(Color.FromArgb(255, 232, 103)))
                    g.FillEllipse(badge, 5 * u, 18 * u, 11 * u, 11 * u);
                using (var mark = new Pen(Color.FromArgb(91, 59, 157), 1.8f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    g.DrawLine(mark, 8 * u, 23.5f * u, 13 * u, 23.5f * u);
                    g.DrawLine(mark, 10.5f * u, 21 * u, 10.5f * u, 26 * u);
                }
            });
        }

        private static Bitmap DrawOpen(int size)
        {
            return Canvas(size, Color.FromArgb(0, 153, 181), Color.FromArgb(0, 111, 166), (g, u) =>
            {
                using (var back = new SolidBrush(Color.FromArgb(190, 239, 250)))
                    FillRound(g, back, 6 * u, 8 * u, 19 * u, 16 * u, 3 * u);
                using (var front = new SolidBrush(Color.White))
                {
                    PointF[] folder =
                    {
                        new PointF(6 * u, 13 * u), new PointF(14 * u, 13 * u),
                        new PointF(16 * u, 10 * u), new PointF(27 * u, 10 * u),
                        new PointF(24 * u, 25 * u), new PointF(7 * u, 25 * u),
                    };
                    g.FillPolygon(front, folder);
                }
                using (var pen = new Pen(Color.FromArgb(0, 111, 166), 2.2f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    g.DrawLine(pen, 17 * u, 15 * u, 17 * u, 22 * u);
                    g.DrawLine(pen, 13.5f * u, 19 * u, 17 * u, 22.5f * u);
                    g.DrawLine(pen, 20.5f * u, 19 * u, 17 * u, 22.5f * u);
                }
            });
        }

        private static Bitmap DrawSave(int size)
        {
            return Canvas(size, Color.FromArgb(0, 166, 118), Color.FromArgb(0, 122, 98), (g, u) =>
            {
                using (var white = new SolidBrush(Color.White))
                    FillRound(g, white, 7 * u, 5.5f * u, 18 * u, 21 * u, 3 * u);
                using (var pale = new SolidBrush(Color.FromArgb(209, 247, 235)))
                    g.FillRectangle(pale, 10 * u, 8 * u, 9 * u, 5 * u);
                using (var pen = new Pen(Color.FromArgb(0, 133, 99), 2.2f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    g.DrawLine(pen, 16 * u, 14 * u, 16 * u, 21 * u);
                    g.DrawLine(pen, 12.5f * u, 18 * u, 16 * u, 21.5f * u);
                    g.DrawLine(pen, 19.5f * u, 18 * u, 16 * u, 21.5f * u);
                    g.DrawLine(pen, 11 * u, 23.5f * u, 21 * u, 23.5f * u);
                }
            });
        }

        private static Bitmap DrawValues(int size)
        {
            return Canvas(size, Color.FromArgb(255, 132, 45), Color.FromArgb(225, 79, 70), (g, u) =>
            {
                using (var fontFx = new Font("Segoe UI", 6.4f * u, FontStyle.Italic | FontStyle.Bold, GraphicsUnit.Pixel))
                using (var fontNum = new Font("Segoe UI", 7.2f * u, FontStyle.Bold, GraphicsUnit.Pixel))
                using (var brush = new SolidBrush(Color.White))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                {
                    g.DrawString("fx", fontFx, brush, new RectangleF(4 * u, 4 * u, 24 * u, 10 * u), format);
                    g.DrawString("123", fontNum, brush, new RectangleF(4 * u, 18 * u, 24 * u, 9 * u), format);
                }
                using (var pen = new Pen(Color.FromArgb(210, 255, 255, 255), 1.6f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    g.DrawLine(pen, 16 * u, 13 * u, 16 * u, 17 * u);
                    g.DrawLine(pen, 13.5f * u, 15.5f * u, 16 * u, 18 * u);
                    g.DrawLine(pen, 18.5f * u, 15.5f * u, 16 * u, 18 * u);
                }
            });
        }

        private static Bitmap DrawSaveAll(int size)
        {
            return Canvas(size, Color.FromArgb(41, 128, 224), Color.FromArgb(52, 86, 190), (g, u) =>
            {
                using (var pale = new SolidBrush(Color.FromArgb(170, 222, 248)))
                    FillRound(g, pale, 6 * u, 6 * u, 13 * u, 17 * u, 2.4f * u);
                using (var white = new SolidBrush(Color.White))
                    FillRound(g, white, 12 * u, 9 * u, 14 * u, 17 * u, 2.4f * u);
                using (var pen = new Pen(Color.FromArgb(42, 105, 195), 2.2f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    g.DrawLine(pen, 15.5f * u, 19 * u, 19 * u, 22.5f * u);
                    g.DrawLine(pen, 19 * u, 22.5f * u, 24 * u, 16.5f * u);
                }
            });
        }

        private static Bitmap DrawIndex(int size)
        {
            return Canvas(size, Color.FromArgb(116, 82, 204), Color.FromArgb(73, 75, 170), (g, u) =>
            {
                using (var white = new SolidBrush(Color.White))
                    FillRound(g, white, 6.5f * u, 5.5f * u, 19 * u, 21 * u, 3 * u);
                using (var dot = new SolidBrush(Color.FromArgb(104, 82, 190)))
                using (var line = new Pen(Color.FromArgb(104, 82, 190), 1.7f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    for (int row = 0; row < 3; row++)
                    {
                        float y = (11 + row * 5) * u;
                        g.FillEllipse(dot, 9 * u, y - 1.2f * u, 2.4f * u, 2.4f * u);
                        g.DrawLine(line, 14 * u, y, 22 * u, y);
                    }
                }
            });
        }

        private static Bitmap DrawTimeline(int size)
        {
            return Canvas(size, Color.FromArgb(221, 70, 120), Color.FromArgb(139, 65, 170), (g, u) =>
            {
                using (var pen = new Pen(Color.White, 2.2f * u)
                { StartCap = LineCap.Round, EndCap = LineCap.Round })
                using (var fill = new SolidBrush(Color.White))
                {
                    g.DrawLine(pen, 10 * u, 8 * u, 10 * u, 24 * u);
                    float[] ys = { 9, 16, 23 };
                    foreach (float y in ys)
                    {
                        g.FillEllipse(fill, 7.5f * u, (y - 2.5f) * u, 5 * u, 5 * u);
                        g.DrawLine(pen, 13 * u, y * u, 24 * u, y * u);
                    }
                }
            });
        }

        private static void FillRound(Graphics g, Brush brush, float x, float y, float width, float height, float radius)
        {
            using (GraphicsPath path = RoundedRect(new RectangleF(x, y, width, height), radius))
                g.FillPath(brush, path);
        }

        private static GraphicsPath RoundedRect(RectangleF bounds, float radius)
        {
            float d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
