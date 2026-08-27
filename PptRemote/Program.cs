using System.Threading;

namespace PptRemote;

internal static class Program
{
    private const string MutexName = @"Local\PptRemote.instance";
    private const string ShowName = @"Local\PptRemote.show";
    private static int _stopping;

    [STAThread]
    private static void Main()
    {
        using var mutex = GrabMutex(out var created);
        using var show = new EventWaitHandle(false, EventResetMode.AutoReset, ShowName);
        if (!created)
        {
            show.Set();
            return;
        }

        ApplicationConfiguration.Initialize();
        var port = Ports.Find(8765);
        var ppt = new PowerPointService();
        var hub = new ClientHub();
        var form = new HostForm(port, ppt, hub);
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
        form.AttachTray(tray);
        tray.MouseClick += (_, e) =>
        {
            if (e.Button == MouseButtons.Left)
            {
                form.ShowFromTray();
            }
        };

        var alive = true;
        var wake = new Thread(() =>
        {
            while (alive)
            {
                if (!show.WaitOne(400))
                {
                    continue;
                }

                if (!alive)
                {
                    break;
                }

                try
                {
                    form.BeginInvoke(form.ShowFromTray);
                }
                catch
                {
                }
            }
        })
        {
            IsBackground = true,
            Name = "ppt-wake"
        };
        wake.Start();

        void Stop()
        {
            if (Interlocked.Exchange(ref _stopping, 1) != 0)
            {
                return;
            }

            alive = false;
            try
            {
                show.Set();
            }
            catch
            {
            }

            try
            {
                tray.Visible = false;
            }
            catch
            {
            }

            try
            {
                Task.Run(() =>
                {
                    try
                    {
                        server.StopAsync().GetAwaiter().GetResult();
                    }
                    catch
                    {
                    }

                    try
                    {
                        ppt.Dispose();
                    }
                    catch
                    {
                    }
                }).Wait(TimeSpan.FromMilliseconds(800));
            }
            catch
            {
            }

            try
            {
                Application.Exit();
            }
            catch
            {
            }

            Environment.Exit(0);
        }

        form.Exiting = Stop;
        form.FormClosed += (_, _) => Stop();
        Application.Run();
        Stop();
    }

    private static Mutex GrabMutex(out bool created)
    {
        try
        {
            return new Mutex(true, MutexName, out created);
        }
        catch (AbandonedMutexException ex)
        {
            created = true;
            return ex.Mutex ?? new Mutex(true, MutexName, out created);
        }
    }
}
