using QRCoder;

namespace PptRemote;

internal sealed class HostForm : Form
{
    private readonly int _port;
    private readonly PowerPointService _ppt;
    private readonly ClientHub _hub;
    private readonly Label _url;
    private readonly RoundBox _deck;
    private readonly Label _deckKick;
    private readonly Label _deckTitle;
    private readonly Label _deckHint;
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
        Size = new Size(328, 560);
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

        _deck = new RoundBox
        {
            Height = 62,
            Dock = DockStyle.Top,
            Fill = Theme.Card,
            Radius = Theme.Radius,
            Margin = new Padding(0, 0, 0, 10)
        };
        _deckKick = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = Theme.Muted,
            Location = new Point(14, 8),
            BackColor = Theme.Card
        };
        _deckTitle = new Label
        {
            AutoSize = false,
            AutoEllipsis = true,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = Theme.Ink,
            Location = new Point(14, 26),
            Size = new Size(250, 22),
            BackColor = Theme.Card
        };
        _deckHint = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 8),
            ForeColor = Theme.Muted,
            Location = new Point(14, 48),
            Visible = false,
            BackColor = Theme.Card
        };
        _deck.Controls.Add(_deckKick);
        _deck.Controls.Add(_deckTitle);
        _deck.Controls.Add(_deckHint);
        _deck.Resize += (_, _) =>
        {
            _deckTitle.Width = Math.Max(40, _deck.ClientSize.Width - 28);
        };
        root.Controls.Add(_deck, 0, 1);

        _qrPane = new Panel { Dock = DockStyle.Fill, BackColor = Theme.Bg };
        _qrWrap = new RoundBox
        {
            Size = new Size(268, 268),
            Fill = Color.White,
            Radius = Theme.Radius
        };
        _qr = new PictureBox
        {
            Size = new Size(248, 248),
            Location = new Point(10, 10),
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
        var urlFont = new Font("Segoe UI", 11);
        var copySize = TextRenderer.MeasureText("Hg", urlFont).Height;
        _url = new Label
        {
            AutoSize = true,
            Font = urlFont,
            ForeColor = Theme.Accent,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 8, 0),
            MaximumSize = new Size(240, 0)
        };
        _url.Click += (_, _) => CopyUrl();
        _copyBtn = new GlyphButton(Glyph.Copy, CopyUrl, copySize);
        _copyBtn.Margin = new Padding(0);
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
        var h = _qrPane.ClientSize.Height;
        var urlH = Math.Max(28, _urlRow.PreferredSize.Height + 6);
        var side = Math.Min(w, Math.Max(160, h - urlH - 8));
        _qrWrap.Size = new Size(side, side);
        var pad = Math.Max(8, side / 26);
        _qr.Location = new Point(pad, pad);
        _qr.Size = new Size(Math.Max(16, side - pad * 2), Math.Max(16, side - pad * 2));
        _qrWrap.Left = Math.Max(0, (w - side) / 2);
        _qrWrap.Top = 0;
        _urlRow.Top = _qrWrap.Bottom + 8;
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
        var key = string.Join("|", addrs.Select(a => a.Name + "=" + a.Ip + "=" + a.Network)) + "#" + _selectedIp;
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
            _statusLine = s.message.Length == 0 ? "Waiting for PowerPoint" : s.message;
            PaintDeck("Waiting", _statusLine, "", Theme.Muted, Theme.Card);
            PushTray();
            return;
        }

        var name = DeckName(s.title);
        if (s.slideshow)
        {
            var hint = $"{s.index} / {s.total}" + (s.black ? "  ·  black" : "");
            _statusLine = $"LIVE  ·  {name}  ·  {hint}";
            PaintDeck("LIVE", name, hint, Theme.Live, Theme.Pin);
            PushTray();
            return;
        }

        _statusLine = $"Deck open  ·  {name}";
        PaintDeck("Deck open", name, "Start from the phone", Theme.Amber, Theme.Wait);
        PushTray();
    }

    private void PaintDeck(string kick, string title, string hint, Color ink, Color fill)
    {
        _deck.Fill = fill;
        _deck.Height = string.IsNullOrEmpty(hint) ? 56 : 70;
        _deckKick.Text = kick;
        _deckKick.ForeColor = ink;
        _deckKick.BackColor = fill;
        _deckTitle.Text = title;
        _deckTitle.ForeColor = ink;
        _deckTitle.BackColor = fill;
        _deckTitle.Width = Math.Max(40, _deck.ClientSize.Width - 28);
        _deckHint.Text = hint;
        _deckHint.ForeColor = ink;
        _deckHint.BackColor = fill;
        _deckHint.Visible = hint.Length > 0;
        _deck.Invalidate();
    }

    private static string DeckName(string title)
    {
        if (string.IsNullOrWhiteSpace(title))
        {
            return "Untitled";
        }

        try
        {
            return Path.GetFileNameWithoutExtension(title);
        }
        catch
        {
            return title;
        }
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
        var network = item.Network.Length == 0 ? item.Name : item.Network;
        var kind = string.Equals(network, item.Name, StringComparison.OrdinalIgnoreCase) ? "" : item.Name;
        return MakeTile(
            pinned ? "This connection" : "Other network",
            network,
            kind,
            pinned,
            pinned ? Theme.Live : Theme.Ink,
            pinned ? null : () => Pick(ip),
            false);
    }

    private Control DeviceTile(string who, string detail, bool phone)
    {
        return MakeTile("Device", who, detail, false, phone ? Theme.Live : Theme.Ink, null, false);
    }

    private Control MakeTile(string kicker, string title, string detail, bool pinned, Color titleColor, Action? click, bool copy)
    {
        var tile = new RoundBox
        {
            Height = string.IsNullOrEmpty(detail) ? 56 : 72,
            Fill = pinned ? Theme.Pin : Theme.Card,
            Radius = Theme.Radius,
            Margin = new Padding(0, 0, 0, 8),
            Padding = new Padding(14, 8, 14, 6),
            Cursor = click == null ? Cursors.Default : Cursors.Hand
        };
        var kick = new Label
        {
            Text = kicker,
            AutoSize = true,
            Font = new Font("Segoe UI", 8, FontStyle.Bold),
            ForeColor = pinned ? Theme.Live : Theme.Muted,
            Location = new Point(14, 8),
            BackColor = pinned ? Theme.Pin : Theme.Card
        };
        var row = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Location = new Point(10, 24),
            BackColor = pinned ? Theme.Pin : Theme.Card,
            Margin = new Padding(0)
        };
        var name = new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI", 11, FontStyle.Bold),
            ForeColor = titleColor,
            Margin = new Padding(4, 4, 8, 0),
            BackColor = pinned ? Theme.Pin : Theme.Card,
            Cursor = click == null && copy ? Cursors.Hand : Cursors.Default
        };
        row.Controls.Add(name);
        if (copy)
        {
            var copyH = TextRenderer.MeasureText("Hg", name.Font).Height;
            var copyBtn = new GlyphButton(Glyph.Copy, CopyUrl, copyH);
            copyBtn.Margin = new Padding(0, 2, 0, 0);
            row.Controls.Add(copyBtn);
            name.Click += (_, _) => CopyUrl();
        }

        tile.Controls.Add(kick);
        tile.Controls.Add(row);
        if (!string.IsNullOrEmpty(detail))
        {
            var line = new Label
            {
                Text = detail,
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = pinned ? Theme.Live : Theme.Muted,
                Location = new Point(14, 54),
                BackColor = pinned ? Theme.Pin : Theme.Card
            };
            tile.Controls.Add(line);
            if (click != null)
            {
                line.Click += (_, _) => click();
            }
        }

        if (click != null)
        {
            void Go(object? s, EventArgs e) => click();
            tile.Click += Go;
            kick.Click += Go;
            name.Click += Go;
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
