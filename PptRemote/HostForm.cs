using QRCoder;

namespace PptRemote;

internal sealed class HostForm : Form
{
    private readonly int _port;
    private readonly PowerPointService _ppt;
    private readonly ClientHub _hub;
    private readonly Label _url;
    private readonly Label _status;
    private readonly VFlow _phones;
    private readonly VFlow _nics;
    private readonly VFlow _linkHost;
    private readonly PictureBox _qr;
    private readonly RoundBox _qrWrap;
    private readonly Panel _qrPane;
    private readonly Panel _devPane;
    private readonly RoundButton _tabQr;
    private readonly RoundButton _tabDev;
    private readonly GlyphButton _copyBtn;
    private readonly FlowLayoutPanel _urlRow;
    private readonly System.Windows.Forms.Timer _tick;
    private NotifyIcon? _tray;
    private string _selectedIp = "";
    private string _nicKey = "";
    private string _viewerKey = "";
    private string _statusKey = "";
    private string _statusLine = "Waiting for PowerPoint";
    private bool _quit;
    private bool _ready;
    private Image? _qrImage;

    public HostForm(int port, PowerPointService ppt, ClientHub hub)
    {
        _port = port;
        _ppt = ppt;
        _hub = hub;
        Text = "PPT Remote";
        Icon = AppIcon.Create();
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        Size = new Size(328, 536);
        BackColor = Theme.Bg;
        ForeColor = Theme.Ink;
        Font = new Font("Segoe UI", 10);
        TopMost = true;
        DoubleBuffered = true;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(18, 16, 18, 18),
            ColumnCount = 1,
            RowCount = 5,
            BackColor = Theme.Bg
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var head = new TableLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 2,
            BackColor = Theme.Bg,
            Margin = new Padding(0, 0, 0, 6)
        };
        head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        head.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 32));
        head.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        head.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        var title = new Label
        {
            Text = "PPT Remote",
            AutoSize = true,
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            ForeColor = Color.White,
            Margin = new Padding(0)
        };
        var sub = new Label
        {
            Text = "Phone presenter · same Wi‑Fi",
            AutoSize = true,
            Font = new Font("Segoe UI", 9),
            ForeColor = Theme.Muted,
            Margin = new Padding(0, 2, 0, 0)
        };
        var close = new GlyphButton(Glyph.Close, HideToTray);
        close.Margin = new Padding(0);
        close.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        head.SetRowSpan(close, 2);
        head.Controls.Add(title, 0, 0);
        head.Controls.Add(sub, 0, 1);
        head.Controls.Add(close, 1, 0);
        root.Controls.Add(head, 0, 0);

        _status = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9, FontStyle.Bold),
            ForeColor = Theme.Muted,
            Margin = new Padding(0, 0, 0, 12),
            MaximumSize = new Size(270, 0)
        };
        root.Controls.Add(_status, 0, 1);

        _qrPane = new Panel { Dock = DockStyle.Fill, BackColor = Theme.Bg };
        _qrWrap = new RoundBox
        {
            Size = new Size(196, 196),
            Fill = Color.White,
            Radius = Theme.Radius
        };
        _qr = new PictureBox
        {
            Size = new Size(172, 172),
            Location = new Point(12, 12),
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White
        };
        _qrWrap.Controls.Add(_qr);
        _urlRow = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = Theme.Bg,
            Padding = new Padding(0),
            Margin = new Padding(0)
        };
        _url = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9),
            ForeColor = Theme.Accent,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 6, 8, 0),
            MaximumSize = new Size(230, 0)
        };
        _url.Click += (_, _) => CopyUrl();
        _copyBtn = new GlyphButton(Glyph.Copy, CopyUrl);
        _copyBtn.Margin = new Padding(0, 2, 0, 0);
        _urlRow.Controls.Add(_url);
        _urlRow.Controls.Add(_copyBtn);
        _qrPane.Controls.Add(_qrWrap);
        _qrPane.Controls.Add(_urlRow);
        _qrPane.Resize += (_, _) => LayoutQr();

        _devPane = new Panel { Dock = DockStyle.Fill, BackColor = Theme.Bg, Visible = false };
        var dev = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            BackColor = Theme.Bg,
            Padding = new Padding(0)
        };
        dev.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        dev.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        dev.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        _linkHost = new VFlow { AutoSize = true, Dock = DockStyle.Top, BackColor = Theme.Bg, Margin = new Padding(0) };
        _nics = new VFlow { AutoSize = true, Dock = DockStyle.Top, BackColor = Theme.Bg, Margin = new Padding(0) };
        _phones = new VFlow { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Theme.Bg, Margin = new Padding(0) };
        dev.Controls.Add(_linkHost, 0, 0);
        dev.Controls.Add(_nics, 0, 1);
        dev.Controls.Add(_phones, 0, 2);
        _devPane.Controls.Add(dev);

        var stage = new Panel { Dock = DockStyle.Fill, BackColor = Theme.Bg };
        stage.Controls.Add(_qrPane);
        stage.Controls.Add(_devPane);
        root.Controls.Add(stage, 0, 2);

        var seg = new RoundBox
        {
            Height = 44,
            Dock = DockStyle.Top,
            Fill = Theme.Card,
            Radius = Theme.Radius,
            Margin = new Padding(0, 12, 0, 12),
            Padding = new Padding(4)
        };
        var segGrid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            BackColor = Color.Transparent,
            Padding = new Padding(0)
        };
        segGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        segGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        _tabQr = SegBtn("QR", true);
        _tabDev = SegBtn("Devices", false);
        _tabQr.Click += (_, _) => SetMode(false);
        _tabDev.Click += (_, _) => SetMode(true);
        segGrid.Controls.Add(_tabQr, 0, 0);
        segGrid.Controls.Add(_tabDev, 1, 0);
        seg.Controls.Add(segGrid);
        root.Controls.Add(seg, 0, 3);

        var quit = new RoundButton
        {
            Text = "Quit",
            Height = 42,
            Dock = DockStyle.Top,
            BackColor = Theme.Accent,
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            Radius = Theme.Radius,
            Cursor = Cursors.Hand,
            Margin = new Padding(0)
        };
        quit.Click += (_, _) => RequestQuit();
        root.Controls.Add(quit, 0, 4);

        _tick = new System.Windows.Forms.Timer { Interval = 400 };
        _tick.Tick += (_, _) =>
        {
            PaintStatus();
            PaintViewers();
        };
        Load += (_, _) =>
        {
            WinChrome.Round(this);
            RefreshNics();
            PaintStatus();
            SetMode(false);
            LayoutQr();
            _tick.Start();
            if (!_ready)
            {
                BeginInvoke(Hide);
            }
        };
        Resize += (_, _) => WinChrome.Round(this);
        var nicTimer = new System.Windows.Forms.Timer { Interval = 4000 };
        nicTimer.Tick += (_, _) => RefreshNics();
        nicTimer.Start();
        FormClosing += OnClosing;
        Deactivate += (_, _) =>
        {
            if (!_quit && Visible)
            {
                Hide();
            }
        };
        Paint += (_, e) =>
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using var pen = new Pen(Theme.Line, 1);
            var r = ClientRectangle;
            r.Width -= 1;
            r.Height -= 1;
            using var path = Theme.Round(r, 14);
            e.Graphics.DrawPath(pen, path);
        };
    }

    public void AttachTray(NotifyIcon tray)
    {
        _tray = tray;
        PushTray();
    }

    public void RequestQuit()
    {
        _quit = true;
        Close();
    }

    public void ShowFromTray()
    {
        _ready = true;
        PlaceNearTray();
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
        LayoutQr();
    }

    private void HideToTray() => Hide();

    private void PlaceNearTray()
    {
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 800, 600);
        Left = wa.Right - Width - 12;
        Top = wa.Bottom - Height - 12;
    }

    private void LayoutQr()
    {
        var w = _qrPane.ClientSize.Width;
        _qrWrap.Left = Math.Max(0, (w - _qrWrap.Width) / 2);
        _qrWrap.Top = 4;
        _urlRow.Top = _qrWrap.Bottom + 10;
        _urlRow.Left = Math.Max(0, (w - _urlRow.Width) / 2);
    }

    private void SetMode(bool devices)
    {
        _qrPane.Visible = !devices;
        _devPane.Visible = devices;
        StyleTab(_tabQr, !devices);
        StyleTab(_tabDev, devices);
        if (devices)
        {
            PaintViewers();
        }
        else
        {
            BeginInvoke(LayoutQr);
        }
    }

    private static void StyleTab(RoundButton b, bool on)
    {
        b.BackColor = on ? Theme.SegOn : Theme.Card;
        b.ForeColor = on ? Theme.Ink : Theme.Muted;
        b.Invalidate();
    }

    private void OnClosing(object? sender, FormClosingEventArgs e)
    {
        if (_quit || e.CloseReason != CloseReason.UserClosing)
        {
            _tick.Stop();
            return;
        }

        e.Cancel = true;
        Hide();
    }

    private void CopyUrl()
    {
        var url = CurrentUrl();
        if (url.Length == 0)
        {
            return;
        }

        Clipboard.SetText(url);
        _url.ForeColor = Theme.Live;
        _copyBtn.IconColor = Theme.Live;
        _copyBtn.Invalidate();
        var t = new System.Windows.Forms.Timer { Interval = 900 };
        t.Tick += (_, _) =>
        {
            _url.ForeColor = Theme.Accent;
            _copyBtn.IconColor = Theme.Ink;
            _copyBtn.Invalidate();
            t.Stop();
            t.Dispose();
        };
        t.Start();
    }

    private string CurrentUrl() => _selectedIp.Length == 0 ? "" : $"http://{_selectedIp}:{_port}";

    private void RefreshNics()
    {
        var addrs = NetworkInfo.ListIpv4();
        var key = string.Join("|", addrs.Select(a => a.Name + "=" + a.Ip)) + "#" + _selectedIp;
        if (key == _nicKey && addrs.Any(a => a.Ip == _selectedIp))
        {
            return;
        }

        _nicKey = key;
        _linkHost.Controls.Clear();
        _nics.Controls.Clear();
        if (addrs.Count == 0)
        {
            _selectedIp = "";
            _url.Text = "No network";
            SetQr("");
            LayoutQr();
            return;
        }

        if (!addrs.Any(a => a.Ip == _selectedIp))
        {
            _selectedIp = addrs[0].Ip;
        }

        ApplyIp(_selectedIp);
        var picked = addrs.First(a => a.Ip == _selectedIp);
        _linkHost.Controls.Add(NicTile(picked, true));
        foreach (var item in addrs.Where(a => a.Ip != _selectedIp))
        {
            _nics.Controls.Add(NicTile(item, false));
        }
    }

    private void Pick(string ip)
    {
        _selectedIp = ip;
        _nicKey = "";
        RefreshNics();
    }

    private void ApplyIp(string ip)
    {
        var url = $"http://{ip}:{_port}";
        _url.Text = url;
        SetQr(url);
        LayoutQr();
    }

    private void SetQr(string url)
    {
        if (url.Length == 0)
        {
            _qr.Image = null;
            return;
        }

        using var gen = new QRCodeGenerator();
        using var data = gen.CreateQrCode(url, QRCodeGenerator.ECCLevel.Q);
        var png = new PngByteQRCode(data);
        using var ms = new MemoryStream(png.GetGraphic(8));
        _qrImage?.Dispose();
        _qrImage = new Bitmap(ms);
        _qr.Image = _qrImage;
    }

    private void PaintStatus()
    {
        var s = _ppt.Snapshot();
        var key = $"{s.connected}|{s.slideshow}|{s.title}|{s.index}|{s.total}|{s.black}|{s.message}";
        if (key == _statusKey)
        {
            return;
        }

        _statusKey = key;
        if (!s.connected)
        {
            _status.ForeColor = Theme.Muted;
            _statusLine = s.message.Length == 0 ? "Waiting for PowerPoint" : s.message;
            _status.Text = _statusLine;
            PushTray();
            return;
        }

        if (s.slideshow)
        {
            _status.ForeColor = Theme.Live;
            _statusLine = $"LIVE  ·  {s.title}  ·  {s.index}/{s.total}" + (s.black ? "  ·  black" : "");
            _status.Text = _statusLine;
            PushTray();
            return;
        }

        _status.ForeColor = Color.FromArgb(251, 191, 36);
        _statusLine = $"Deck open  ·  {s.title}  ·  start from the phone";
        _status.Text = _statusLine;
        PushTray();
    }

    private void PushTray()
    {
        if (_tray == null)
        {
            return;
        }

        var tip = "PPT Remote — " + _statusLine;
        _tray.Text = tip.Length <= 63 ? tip : tip[..60] + "...";
    }

    private void PaintViewers()
    {
        var list = _hub.Active();
        var key = string.Join("|", list.Select(v => v.Ip + "@" + ((int)(DateTime.Now - v.LastSeen).TotalSeconds)));
        if (key == _viewerKey)
        {
            return;
        }

        _viewerKey = key;
        _phones.Controls.Clear();
        if (list.Count == 0)
        {
            _phones.Controls.Add(MutedLine("Nobody connected yet."));
            return;
        }

        _phones.Controls.Add(MutedLine("Connected"));
        foreach (var v in list)
        {
            var age = Math.Max(0, (int)(DateTime.Now - v.LastSeen).TotalSeconds);
            var who = v.IsLocal ? "This PC" : "Phone";
            var live = age <= 2 ? "live" : age + "s ago";
            _phones.Controls.Add(DeviceTile(who, v.Ip + "  ·  " + live, !v.IsLocal));
        }
    }

    private Control NicTile(NicAddr item, bool pinned)
    {
        var ip = item.Ip;
        return MakeTile(
            pinned ? "This connection" : "Other network",
            item.Name,
            ip,
            pinned,
            pinned ? Theme.Accent : Theme.Ink,
            pinned ? null : () => Pick(ip),
            pinned);
    }

    private Control DeviceTile(string who, string detail, bool phone)
    {
        return MakeTile("Device", who, detail, false, phone ? Theme.Live : Theme.Ink, null, false);
    }

    private Control MakeTile(string kicker, string title, string detail, bool pinned, Color titleColor, Action? click, bool copy)
    {
        var tile = new RoundBox
        {
            Height = 86,
            Fill = pinned ? Theme.Pin : Theme.Card,
            Radius = Theme.Radius,
            Margin = new Padding(0, 0, 0, 8),
            Padding = new Padding(14, 10, 14, 10),
            Cursor = click == null ? Cursors.Default : Cursors.Hand
        };
        var kick = new Label
        {
            Text = kicker,
            AutoSize = true,
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = pinned ? Theme.Accent : Theme.Muted,
            Location = new Point(14, 10),
            BackColor = pinned ? Theme.Pin : Theme.Card
        };
        var name = new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = titleColor,
            Location = new Point(14, 28),
            BackColor = pinned ? Theme.Pin : Theme.Card
        };
        var row = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Location = new Point(10, 50),
            BackColor = pinned ? Theme.Pin : Theme.Card,
            Margin = new Padding(0)
        };
        var line = new Label
        {
            Text = detail,
            AutoSize = true,
            Font = new Font("Segoe UI", 9),
            ForeColor = pinned ? Theme.Accent : Theme.Muted,
            Margin = new Padding(4, 6, 8, 0),
            BackColor = pinned ? Theme.Pin : Theme.Card,
            Cursor = Cursors.Hand
        };
        row.Controls.Add(line);
        if (copy)
        {
            var copyBtn = new GlyphButton(Glyph.Copy, CopyUrl);
            copyBtn.Margin = new Padding(0, 0, 0, 0);
            row.Controls.Add(copyBtn);
        }

        tile.Controls.Add(kick);
        tile.Controls.Add(name);
        tile.Controls.Add(row);
        if (click != null)
        {
            void Go(object? s, EventArgs e) => click();
            tile.Click += Go;
            kick.Click += Go;
            name.Click += Go;
            line.Click += Go;
        }
        else if (copy)
        {
            line.Click += (_, _) => CopyUrl();
        }

        return tile;
    }

    private static Label MutedLine(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = Theme.Muted,
            Margin = new Padding(0, 4, 0, 8),
            BackColor = Color.Transparent
        };
    }

    private static RoundButton SegBtn(string text, bool on)
    {
        var b = new RoundButton
        {
            Text = text,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 10, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(2),
            Radius = 8
        };
        StyleTab(b, on);
        return b;
    }
}
