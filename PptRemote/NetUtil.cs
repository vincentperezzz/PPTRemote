using System.Diagnostics;
using System.Net;
using System.Net.Sockets;

namespace PptRemote;

internal static class Ports
{
    public static int Find(int preferred)
    {
        for (var port = preferred; port < preferred + 12; port++)
        {
            var listener = new TcpListener(IPAddress.Any, port);
            try
            {
                listener.Start();
                listener.Stop();
                return port;
            }
            catch (SocketException)
            {
            }
            finally
            {
                listener.Stop();
            }
        }

        throw new InvalidOperationException("No free port");
    }
}

internal static class FirewallHelper
{
    public static bool TryAllow(int port, string exePath)
    {
        TryDelete("PPT Remote");
        TryDelete("PPT Remote Port");
        var okProg = Run($"advfirewall firewall add rule name=\"PPT Remote\" dir=in action=allow program=\"{exePath}\" enable=yes profile=any");
        var okPort = Run($"advfirewall firewall add rule name=\"PPT Remote Port\" dir=in action=allow protocol=TCP localport={port} profile=any");
        return okProg || okPort;
    }

    public static void PromptAllow(int port, string exePath)
    {
        var bat = Path.Combine(Path.GetTempPath(), "ppt-remote-fw.bat");
        File.WriteAllText(bat,
            "@echo off\r\n" +
            "netsh advfirewall firewall delete rule name=\"PPT Remote\" >nul 2>&1\r\n" +
            "netsh advfirewall firewall delete rule name=\"PPT Remote Port\" >nul 2>&1\r\n" +
            "netsh advfirewall firewall add rule name=\"PPT Remote\" dir=in action=allow program=\"" + exePath + "\" enable=yes profile=any\r\n" +
            "netsh advfirewall firewall add rule name=\"PPT Remote Port\" dir=in action=allow protocol=TCP localport=" + port + " profile=any\r\n");
        try
        {
            using var p = Process.Start(new ProcessStartInfo
            {
                FileName = bat,
                Verb = "runas",
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            p?.WaitForExit(20000);
        }
        catch
        {
        }
    }

    private static void TryDelete(string name)
    {
        Run($"advfirewall firewall delete rule name=\"{name}\"");
    }

    private static bool Run(string args)
    {
        try
        {
            using var p = Process.Start(new ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = args,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            if (p == null)
            {
                return false;
            }

            p.WaitForExit(8000);
            return p.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }
}
