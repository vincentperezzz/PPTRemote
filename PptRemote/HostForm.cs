using QRCoder;

namespace PptRemote;

internal sealed class HostForm : Form
{
    private readonly int _port;
    private readonly PowerPointService _ppt;
    private readonly ClientHub _hub;
    private readonly Label _ip;
    private readonly Label _url;
    private readonly Label _status;
    private readonly Label _viewers;
    private readonly Label _hint;
    private readonly PictureBox _qr;
    private readonly FlowLayoutPanel _nics;
    private readonly System.Windows.Forms.Timer _tick;
    private string _selectedIp = "";
    private string _nicKey = "";
    private string _viewerKey = "";
    private bool _quit;
    private Image? _qrImage;

    public HostForm(int port, PowerPointService ppt, ClientHub hub)
    {
        _port = port;
        _ppt = ppt;
        _hub = hub;
        Text = "PPT Remote";
        Icon = AppIcon.Create();
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(420, 740);
        Size = new Size(460, 820);
        BackColor = Color.FromArgb(12, 12, 14);
        ForeColor = Color.FromArgb(244, 244, 245);
        Font = new Font("Segoe UI", 10);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(22, 18, 22, 16),
            ColumnCount = 1,
            RowCount = 10,
            BackColor = BackColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        root.Controls.Add(MkLabel("PPT Remote", 22, FontStyle.Bold, Color.White), 0, 0);
        root.Controls.Add(MkLabel("On iPhone Safari, type this — or scan. Same Wi-Fi / hotspot.", 9, FontStyle.Regular, Color.FromArgb(154, 154, 163)), 0, 1);

        _ip = MkLabel("—", 26, FontStyle.Bold, Color.White);
        _ip.Font = new Font("Cascadia Mono", 22, FontStyle.Bold);
        root.Controls.Add(_ip, 0, 2);

        _url = MkLabel("", 11, FontStyle.Regular, Color.FromArgb(255, 77, 46));
        root.Controls.Add(_url, 0, 3);

        var btns = new FlowLayoutPanel { AutoSize = true, WrapContents = false, BackColor = BackColor };
        btns.Controls.Add(MkBtn("Copy address", CopyUrl));
        btns.Controls.Add(MkBtn("Allow iPhone", AllowPhone));
        btns.Controls.Add(MkBtn("Hide to tray", HideToTray));
        root.Controls.Add(btns, 0, 4);

        _qr = new PictureBox
        {
            Size = new Size(180, 180),
            SizeMode = PictureBoxSizeMode.Zoom,
            BackColor = Color.FromArgb(22, 22, 26),
            Margin = new Padding(0, 12, 0, 8)
        };
        root.Controls.Add(_qr, 0, 5);

        _viewers = MkLabel("Viewing this URL\nNobody yet — open it on the iPhone.", 10, FontStyle.Regular, Color.FromArgb(212, 212, 216));
        root.Controls.Add(_viewers, 0, 6);

        _nics = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            BackColor = BackColor
        };
        root.Controls.Add(_nics, 0, 7);

        _status = MkLabel("", 10, FontStyle.Regular, Color.FromArgb(154, 154, 163));
        root.Controls.Add(_status, 0, 8);

        _hint = MkLabel("If Safari loads then gets stuck, tap Allow iPhone once (blue Yes on the Windows prompt).", 9, FontStyle.Regular, Color.FromArgb(154, 154, 163));
        root.Controls.Add(_hint, 0, 9);

        var quitRow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, BackColor = BackColor };
        quitRow.Controls.Add(MkBtn("Quit", RequestQuit));
        Controls.Add(quitRow);

        _tick = new System.Windows.Forms.Timer { Interval = 400 };
        _tick.Tick += (_, _) =>
        {
            PaintStatus();
            PaintViewers();
        };
        Shown += (_, _) =>
        {
            RefreshNics();
            _tick.Start();
        };
        var nicTimer = new System.Windows.Forms.Timer { Interval = 4000 };
        nicTimer.Tick += (_, _) => RefreshNics();
        nicTimer.Start();
        FormClosing += OnClosing;
    }

    public void RequestQuit()
    {
        _quit = true;
        Close();
    }

    public void ShowFromTray()
    {
        Show();
        WindowState = FormWindowState.Normal;
        Activate();
    }

    private void HideToTray()
    {
        Hide();
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
        _hint.Text = "Copied. Paste it into Safari.";
    }

    private void AllowPhone()
    {
        _hint.Text = "Windows will ask to allow this app. Tap Yes.";
        Refresh();
        FirewallHelper.PromptAllow(_port, Application.ExecutablePath);
        _hint.Text = "If you tapped Yes, pull-to-refresh Safari. The phone should pop up under Viewing.";
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
            _ip.Text = "No network";
            _url.Text = "Turn on Wi-Fi or a hotspot";
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
            var btn = MkBtn($"{item.Name}   {ip}", () => Pick(ip));
            btn.ForeColor = ip == _selectedIp ? Color.FromArgb(255, 77, 46) : Color.White;
            btn.Width = 360;
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
        var url = $"http://{ip}:{_port}";
        _ip.Text = ip;
        _url.Text = url;
        SetQr(url);
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
        using var ms = new MemoryStream(png.GetGraphic(7));
        _qrImage?.Dispose();
        _qrImage = new Bitmap(ms);
        _qr.Image = _qrImage;
    }

    private void PaintStatus()
    {
        var s = _ppt.Snapshot();
        if (!s.connected)
        {
            _status.ForeColor = Color.FromArgb(154, 154, 163);
            _status.Text = s.message;
            return;
        }

        if (s.slideshow)
        {
            _status.ForeColor = Color.FromArgb(74, 222, 128);
            _status.Text = $"Live  ·  {s.title}  ·  slide {s.index}/{s.total}" + (s.black ? "  ·  black" : "");
            return;
        }

        _status.ForeColor = Color.FromArgb(251, 191, 36);
        _status.Text = $"Deck open  ·  {s.title}  ·  start slideshow from the phone";
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
            _viewers.Text = "Viewing this URL\nNobody yet — open it on the iPhone.";
            _viewers.ForeColor = Color.FromArgb(154, 154, 163);
            return;
        }

        var lines = new List<string> { "Viewing this URL" };
        foreach (var v in list)
        {
            var age = Math.Max(0, (int)(DateTime.Now - v.LastSeen).TotalSeconds);
            var who = v.IsLocal ? "This PC" : "Phone";
            var live = age <= 2 ? "live" : age + "s ago";
            lines.Add($"{who}  ·  {v.Ip}  ·  {live}");
        }

        _viewers.Text = string.Join("\n", lines);
        _viewers.ForeColor = Color.FromArgb(74, 222, 128);
    }

    private static Label MkLabel(string text, float size, FontStyle style, Color color)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            ForeColor = color,
            Font = new Font("Segoe UI", size, style),
            MaximumSize = new Size(380, 0)
        };
    }

    private Button MkBtn(string text, Action click)
    {
        var b = new Button
        {
            Text = text,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(22, 22, 26),
            ForeColor = Color.White,
            Padding = new Padding(10, 6, 10, 6),
            Margin = new Padding(0, 0, 8, 8),
            Cursor = Cursors.Hand
        };
        b.FlatAppearance.BorderColor = Color.FromArgb(39, 39, 44);
        b.Click += (_, _) => click();
        return b;
    }
}
