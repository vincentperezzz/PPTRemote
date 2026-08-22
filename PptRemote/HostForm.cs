using QRCoder;

namespace PptRemote;

internal sealed class HostForm : Form
{
    private readonly int _port;
    private readonly ClientHub _hub;
    private readonly Label _viewers;
    private readonly PictureBox _qr;
    private readonly FlowLayoutPanel _nics;
    private readonly Panel _extra;
    private readonly Button _devicesBtn;
    private readonly System.Windows.Forms.Timer _tick;
    private string _selectedIp = "";
    private string _nicKey = "";
    private string _viewerKey = "";
    private bool _quit;
    private bool _ready;
    private Image? _qrImage;

    public HostForm(int port, ClientHub hub)
    {
        _port = port;
        _hub = hub;
        Text = "PPT Remote";
        Icon = AppIcon.Create();
        ShowInTaskbar = false;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition = FormStartPosition.Manual;
        Size = new Size(280, 420);
        BackColor = Color.FromArgb(18, 18, 20);
        ForeColor = Color.FromArgb(244, 244, 245);
        Font = new Font("Segoe UI", 10);
        TopMost = true;
        Padding = new Padding(0);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(14, 10, 14, 12),
            ColumnCount = 1,
            RowCount = 5,
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var head = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 3,
            RowCount = 1,
            BackColor = BackColor,
            Margin = new Padding(0, 0, 0, 8)
        };
        head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        head.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 32));
        head.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 32));
        var title = new Label
        {
            Text = "PPT Remote",
            AutoSize = true,
            Font = new Font("Segoe UI", 13, FontStyle.Bold),
            ForeColor = Color.White,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft
        };
        head.Controls.Add(title, 0, 0);
        head.Controls.Add(IconBtn("⧉", CopyUrl), 1, 0);
        head.Controls.Add(IconBtn("✕", HideToTray), 2, 0);
        root.Controls.Add(head, 0, 0);

        _qr = new PictureBox
        {
            Size = new Size(220, 220),
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.White,
            Margin = new Padding(16, 4, 16, 8),
            Anchor = AnchorStyles.None
        };
        var qrWrap = new FlowLayoutPanel
        {
            AutoSize = true,
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = BackColor
        };
        qrWrap.Controls.Add(_qr);
        root.Controls.Add(qrWrap, 0, 1);

        _devicesBtn = MkBtn("Devices  ▸", ToggleExtra);
        _devicesBtn.Dock = DockStyle.Top;
        _devicesBtn.Width = 240;
        root.Controls.Add(_devicesBtn, 0, 2);

        _extra = new Panel
        {
            Visible = false,
            AutoSize = true,
            Dock = DockStyle.Fill,
            BackColor = BackColor
        };
        _viewers = new Label
        {
            Text = "Nobody connected.",
            AutoSize = true,
            ForeColor = Color.FromArgb(154, 154, 163),
            Font = new Font("Segoe UI", 9),
            MaximumSize = new Size(240, 0),
            Margin = new Padding(0, 4, 0, 8)
        };
        _nics = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            BackColor = BackColor,
            Dock = DockStyle.Top
        };
        _extra.Controls.Add(_nics);
        _extra.Controls.Add(_viewers);
        _viewers.Dock = DockStyle.Top;
        _nics.Dock = DockStyle.Top;
        root.Controls.Add(_extra, 0, 3);

        var quit = MkBtn("Quit", RequestQuit);
        quit.Dock = DockStyle.Top;
        quit.BackColor = Color.FromArgb(255, 77, 46);
        quit.ForeColor = Color.White;
        quit.FlatAppearance.BorderColor = Color.FromArgb(255, 77, 46);
        root.Controls.Add(quit, 0, 4);

        _tick = new System.Windows.Forms.Timer { Interval = 400 };
        _tick.Tick += (_, _) => PaintViewers();
        Load += (_, _) =>
        {
            RefreshNics();
            _tick.Start();
            if (!_ready)
            {
                BeginInvoke(Hide);
            }
        };
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
    }

    private void HideToTray()
    {
        Hide();
    }

    private void PlaceNearTray()
    {
        var wa = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 800, 600);
        Left = wa.Right - Width - 12;
        Top = wa.Bottom - Height - 12;
    }

    private void ToggleExtra()
    {
        _extra.Visible = !_extra.Visible;
        _devicesBtn.Text = _extra.Visible ? "Devices  ▾" : "Devices  ▸";
        Height = _extra.Visible ? 520 : 420;
        PlaceNearTray();
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
    }

    private string CurrentUrl() => _selectedIp.Length == 0 ? "" : $"http://{_selectedIp}:{_port}";

    private void RefreshNics()
    {
        var addrs = NetworkInfo.ListIpv4();
        var key = string.Join("|", addrs.Select(a => a.Name + "=" + a.Ip));
        if (key == _nicKey && addrs.Any(a => a.Ip == _selectedIp))
        {
            return;
        }

        _nicKey = key;
        if (addrs.Count == 0)
        {
            _selectedIp = "";
            SetQr("");
            _nics.Controls.Clear();
            return;
        }

        if (!addrs.Any(a => a.Ip == _selectedIp))
        {
            _selectedIp = addrs[0].Ip;
        }

        ApplyIp(_selectedIp);
        _nics.Controls.Clear();
        foreach (var item in addrs)
        {
            var ip = item.Ip;
            var btn = MkBtn($"{item.Name}  {ip}", () => Pick(ip));
            btn.ForeColor = ip == _selectedIp ? Color.FromArgb(255, 77, 46) : Color.White;
            btn.Width = 240;
            _nics.Controls.Add(btn);
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
        SetQr($"http://{ip}:{_port}");
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

    private void PaintViewers()
    {
        var list = _hub.Active();
        var key = string.Join("|", list.Select(v => v.Ip + "@" + v.LastSeen.Ticks));
        if (key == _viewerKey)
        {
            return;
        }

        _viewerKey = key;
        if (list.Count == 0)
        {
            _viewers.Text = "Nobody connected.";
            _viewers.ForeColor = Color.FromArgb(154, 154, 163);
            _devicesBtn.Text = _extra.Visible ? "Devices  ▾" : "Devices  ▸";
            return;
        }

        var lines = new List<string>();
        foreach (var v in list)
        {
            var age = Math.Max(0, (int)(DateTime.Now - v.LastSeen).TotalSeconds);
            var who = v.IsLocal ? "This PC" : "Phone";
            var live = age <= 2 ? "live" : age + "s ago";
            lines.Add($"{who}  ·  {v.Ip}  ·  {live}");
        }

        _viewers.Text = string.Join("\n", lines);
        _viewers.ForeColor = Color.FromArgb(74, 222, 128);
        _devicesBtn.Text = (_extra.Visible ? "Devices  ▾  " : "Devices  ▸  ") + list.Count;
    }

    private Button IconBtn(string text, Action click)
    {
        var b = new Button
        {
            Text = text,
            Width = 28,
            Height = 28,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(32, 32, 36),
            ForeColor = Color.White,
            Margin = new Padding(2, 0, 0, 0),
            Cursor = Cursors.Hand,
            Font = new Font("Segoe UI", 10)
        };
        b.FlatAppearance.BorderColor = Color.FromArgb(50, 50, 56);
        b.Click += (_, _) => click();
        return b;
    }

    private Button MkBtn(string text, Action click)
    {
        var b = new Button
        {
            Text = text,
            AutoSize = false,
            Height = 36,
            Width = 240,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(32, 32, 36),
            ForeColor = Color.White,
            Margin = new Padding(0, 0, 0, 8),
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.MiddleCenter,
            Padding = new Padding(10, 0, 10, 0)
        };
        b.FlatAppearance.BorderColor = Color.FromArgb(50, 50, 56);
        b.Click += (_, _) => click();
        return b;
    }
}
