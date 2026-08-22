namespace PptRemote;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        var port = Ports.Find(8765);
        var ppt = new PowerPointService();
        var hub = new ClientHub();
        var form = new HostForm(port, hub);
        _ = form.Handle;
        var server = AppServer.Start(port, ppt, hub);
        FirewallHelper.TryAllow(port, Application.ExecutablePath);

        using var tray = new NotifyIcon
        {
            Icon = form.Icon,
            Text = $"PPT Remote  :{port}",
            Visible = true,
            ContextMenuStrip = new ContextMenuStrip()
        };
        tray.ContextMenuStrip.Items.Add("Open", null, (_, _) => form.ShowFromTray());
        tray.ContextMenuStrip.Items.Add("Quit", null, (_, _) => form.RequestQuit());
        tray.MouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                form.ShowFromTray();
            }
        };

        form.FormClosed += (_, _) =>
        {
            tray.Visible = false;
            try
            {
                server.StopAsync().GetAwaiter().GetResult();
            }
            catch
            {
            }

            ppt.Dispose();
            Application.Exit();
        };

        Application.Run();
    }
}
