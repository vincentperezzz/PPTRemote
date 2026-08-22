namespace PptRemote;

internal static class AppIcon
{
    public static Icon Create()
    {
        var bmp = new Bitmap(32, 32);
        using (var g = Graphics.FromImage(bmp))
        {
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.Clear(Color.Transparent);
            using var brush = new SolidBrush(Color.FromArgb(255, 77, 46));
            g.FillEllipse(brush, 1, 1, 30, 30);
            using var white = new SolidBrush(Color.White);
            var play = new Point[]
            {
                new Point(12, 8),
                new Point(12, 24),
                new Point(24, 16)
            };
            g.FillPolygon(white, play);
        }

        return Icon.FromHandle(bmp.GetHicon());
    }
}
