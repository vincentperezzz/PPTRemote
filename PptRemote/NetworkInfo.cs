using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Net.Sockets;
namespace PptRemote;

internal sealed class NicAddr
{
    public string Name { get; init; } = "";
    public string Ip { get; init; } = "";
    public string Network { get; init; } = "";
}

internal static class NetworkInfo
{
    private static Dictionary<string, string>? _ssidCache;
    private static DateTime _ssidAt;

    public static List<NicAddr> ListIpv4()
    {
        var found = new List<NicAddr>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var ssids = WifiSsids();
        var profiles = NetworkProfiles();
        foreach (var nic in NetworkInterface.GetAllNetworkInterfaces())
        {
            if (nic.OperationalStatus != OperationalStatus.Up)
            {
                continue;
            }

            if (IsVirtual(nic.Name, nic.Description))
            {
                continue;
            }

            foreach (var addr in nic.GetIPProperties().UnicastAddresses)
            {
                if (addr.Address.AddressFamily != AddressFamily.InterNetwork)
                {
                    continue;
                }

                var ip = addr.Address.ToString();
                if (ip.StartsWith("127.") || ip.StartsWith("169.254."))
                {
                    continue;
                }

                if (!seen.Add(ip))
                {
                    continue;
                }

                var network = ResolveName(nic, ip, ssids, profiles);
                found.Add(new NicAddr { Name = nic.Name, Ip = ip, Network = network });
            }
        }

        found.Sort(Compare);
        return found;
    }

    private static string ResolveName(
        NetworkInterface nic,
        string ip,
        Dictionary<string, string> ssids,
        Dictionary<string, string> profiles)
    {
        if (ssids.TryGetValue(nic.Name, out var ssid) && ssid.Length > 0)
        {
            return ssid;
        }

        if (profiles.TryGetValue(nic.Id, out var profile) && profile.Length > 0)
        {
            return profile;
        }

        if (ip.StartsWith("192.168.137."))
        {
            return "Hotspot";
        }

        return nic.Name;
    }

    private static Dictionary<string, string> WifiSsids()
    {
        if (_ssidCache != null && DateTime.UtcNow - _ssidAt < TimeSpan.FromSeconds(8))
        {
            return _ssidCache;
        }

        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "netsh",
                Arguments = "wlan show interfaces",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var proc = Process.Start(psi);
            if (proc == null)
            {
                return map;
            }

            var text = proc.StandardOutput.ReadToEnd();
            if (!proc.WaitForExit(1500))
            {
                try
                {
                    proc.Kill();
                }
                catch
                {
                }

                return map;
            }

            string? adapter = null;
            foreach (var raw in text.Split('\n'))
            {
                var line = raw.Trim();
                var split = line.IndexOf(':');
                if (split < 0)
                {
                    continue;
                }

                var key = line[..split].Trim();
                var value = line[(split + 1)..].Trim();
                if (key.Equals("Name", StringComparison.OrdinalIgnoreCase))
                {
                    adapter = value;
                    continue;
                }

                if (adapter == null || value.Length == 0)
                {
                    continue;
                }

                if (key.Equals("SSID", StringComparison.OrdinalIgnoreCase) ||
                    key.Equals("Profile", StringComparison.OrdinalIgnoreCase))
                {
                    map[adapter] = value;
                }
            }
        }
        catch
        {
        }

        _ssidCache = map;
        _ssidAt = DateTime.UtcNow;
        return map;
    }

    private static Dictionary<string, string> NetworkProfiles()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var type = Type.GetTypeFromCLSID(new Guid("DCB00C01-570F-4A9B-8D69-199FDBA5723B"));
            if (type == null)
            {
                return map;
            }

            dynamic nlm = Activator.CreateInstance(type)!;
            foreach (dynamic conn in nlm.GetNetworkConnections())
            {
                Guid id = conn.GetAdapterId();
                string name = conn.GetNetwork().GetName();
                if (!string.IsNullOrWhiteSpace(name))
                {
                    map[id.ToString("B")] = name;
                }
            }
        }
        catch
        {
        }

        return map;
    }

    private static bool IsVirtual(string name, string desc)
    {
        var s = (name + " " + desc).ToLowerInvariant();
        return s.Contains("vethernet")
            || s.Contains("hyper-v")
            || s.Contains("wsl")
            || s.Contains("virtualbox")
            || s.Contains("vmware")
            || s.Contains("loopback")
            || s.Contains("bluetooth");
    }

    private static int Compare(NicAddr a, NicAddr b)
    {
        var c = Rank(a).CompareTo(Rank(b));
        return c != 0 ? c : string.CompareOrdinal(a.Ip, b.Ip);
    }

    private static int Rank(NicAddr item)
    {
        var name = item.Name.ToLowerInvariant();
        var ip = item.Ip;
        if (ip.StartsWith("192.168.137."))
        {
            return 0;
        }

        if (name.Contains("hosted") || name.Contains("hotspot") || name.Contains("local area connection"))
        {
            return 1;
        }

        if (name.StartsWith("wi-fi") || name.StartsWith("wifi"))
        {
            return 2;
        }

        if (ip.StartsWith("192.168."))
        {
            return 3;
        }

        if (ip.StartsWith("10."))
        {
            return 4;
        }

        return 5;
    }
}
