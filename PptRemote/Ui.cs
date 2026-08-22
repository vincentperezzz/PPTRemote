using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace PptRemote;

internal static class Theme
{
    public const int Radius = 12;
    public static readonly Color Bg = Color.FromArgb(10, 10, 11);
    public static readonly Color Card = Color.FromArgb(22, 22, 24);
    public static readonly Color Line = Color.FromArgb(42, 42, 46);
    public static readonly Color Ink = Color.FromArgb(243, 239, 232);
    public static readonly Color Muted = Color.FromArgb(156, 152, 144);
    public static readonly Color Accent = Color.FromArgb(255, 77, 46);
    public static readonly Color Live = Color.FromArgb(61, 220, 132);
    public static readonly Color SegOn = Color.FromArgb(42, 42, 47);
    public static readonly Color Pin = Color.FromArgb(16, 38, 26);

    public static Color SolidBack(Control? c)
    {
        while (c != null)
        {
            if (c is RoundBox box)
            {
                return box.Fill;
            }

            if (c.BackColor.A == 255 && c.BackColor != Color.Transparent)
            {
                return c.BackColor;
            }

            c = c.Parent;
        }

        return Bg;
    }

    public static GraphicsPath Round(Rectangle r, int radius)
    {
        var d = Math.Max(2, radius * 2);
        if (d > r.Width) d = r.Width;
        if (d > r.Height) d = r.Height;
        var p = new GraphicsPath();
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }
}

internal static class WinChrome
{
    private const int DwmwaWindowCornerPreference = 33;
    private const int DwmwcpRound = 2;

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int val, int size);

    public static void Round(Form form)
    {
        var v = DwmwcpRound;
        try
        {
            _ = DwmSetWindowAttribute(form.Handle, DwmwaWindowCornerPreference, ref v, sizeof(int));
        }
        catch
        {
        }
    }
}

internal sealed class RoundBox : Panel
{
    public Color Fill { get; set; } = Theme.Card;
    public int Radius { get; set; } = Theme.Radius;

    public RoundBox()
    {
        DoubleBuffered = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        BackColor = Theme.Bg;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        using var path = Theme.Round(r, Radius);
        using var br = new SolidBrush(Fill);
        e.Graphics.FillPath(br, path);
    }
}

internal sealed class RoundButton : Button
{
    public int Radius { get; set; } = Theme.Radius;

    public RoundButton()
    {
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        FlatAppearance.MouseOverBackColor = Color.Transparent;
        FlatAppearance.MouseDownBackColor = Color.Transparent;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        e.Graphics.Clear(Theme.SolidBack(Parent));
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        using var path = Theme.Round(r, Radius);
        using var br = new SolidBrush(BackColor);
        e.Graphics.FillPath(br, path);
        TextRenderer.DrawText(
            e.Graphics,
            Text,
            Font,
            ClientRectangle,
            ForeColor,
            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }
}

internal enum Glyph
{
    Close,
    Copy
}

internal sealed class GlyphButton : Control
{
    private readonly Glyph _glyph;
    private bool _hot;
    private bool _down;

    public Color IconColor { get; set; } = Theme.Ink;

    public GlyphButton(Glyph glyph, Action click, int size = 28)
    {
        _glyph = glyph;
        Size = new Size(size, size);
        Cursor = Cursors.Hand;
        DoubleBuffered = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        Click += (_, _) => click();
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        _hot = true;
        Invalidate();
        base.OnMouseEnter(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hot = false;
        _down = false;
        Invalidate();
        base.OnMouseLeave(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        _down = true;
        Invalidate();
        base.OnMouseDown(e);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        _down = false;
        Invalidate();
        base.OnMouseUp(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
        e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        e.Graphics.Clear(Theme.SolidBack(Parent));
        var fill = _down ? Theme.Line : _hot ? Color.FromArgb(36, 36, 40) : Theme.Card;
        var r = new Rectangle(0, 0, Width - 1, Height - 1);
        using (var path = Theme.Round(r, 8))
        using (var br = new SolidBrush(fill))
        {
            e.Graphics.FillPath(br, path);
        }

        var u = Width / 28f;
        using var pen = new Pen(IconColor, Math.Max(1.6f, 1.7f * u))
        {
            StartCap = LineCap.Round,
            EndCap = LineCap.Round,
            LineJoin = LineJoin.Round
        };
        if (_glyph == Glyph.Close)
        {
            var a = 9 * u;
            var b = Width - 10 * u;
            e.Graphics.DrawLine(pen, a, a, b, Height - 10 * u);
            e.Graphics.DrawLine(pen, b, a, a, Height - 10 * u);
            return;
        }

        var back = Rectangle.Round(new RectangleF(8 * u, 7 * u, 11 * u, 13 * u));
        var front = Rectangle.Round(new RectangleF(11 * u, 10 * u, 11 * u, 13 * u));
        using (var path = Theme.Round(back, Math.Max(2, (int)(2 * u))))
        {
            e.Graphics.DrawPath(pen, path);
        }

        using (var path = Theme.Round(front, Math.Max(2, (int)(2 * u))))
        using (var hide = new SolidBrush(fill))
        {
            e.Graphics.FillPath(hide, path);
            e.Graphics.DrawPath(pen, path);
        }
    }
}

internal sealed class VFlow : FlowLayoutPanel
{
    [DllImport("user32.dll")]
    private static extern int ShowScrollBar(IntPtr hwnd, int bar, bool show);

    public VFlow()
    {
        FlowDirection = FlowDirection.TopDown;
        WrapContents = false;
        DoubleBuffered = true;
        AutoScroll = false;
    }

    protected override void OnLayout(LayoutEventArgs levent)
    {
        var inner = ClientSize.Width - Padding.Horizontal;
        if (AutoScroll && VerticalScroll.Visible)
        {
            inner -= SystemInformation.VerticalScrollBarWidth;
        }

        inner = Math.Max(24, inner);
        foreach (Control c in Controls)
        {
            if (c.Dock == DockStyle.None)
            {
                c.Width = Math.Max(24, inner - c.Margin.Horizontal);
            }
        }

        base.OnLayout(levent);
        if (IsHandleCreated)
        {
            ShowScrollBar(Handle, 0, false);
        }
    }
}
