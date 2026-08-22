using System.Net.NetworkInformation;
using System.Net.Sockets;

namespace PptRemote;

internal sealed class NicAddr
{
    public string Name { get; init; } = "";
    public string Ip { get; init; } = "";
}

internal static class NetworkInfo
{
    public static List<NicAddr> ListIpv4()
    {
        var found = new List<NicAddr>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
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

                found.Add(new NicAddr { Name = nic.Name, Ip = ip });
            }
        }

        found.Sort(Compare);
        return found;
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
        var ra = Rank(a);
        var rb = Rank(b);
        var c = ra.CompareTo(rb);
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
