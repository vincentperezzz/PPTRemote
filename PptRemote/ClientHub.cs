namespace PptRemote;

internal sealed class Viewer
{
    public string Ip { get; init; } = "";
    public DateTime LastSeen { get; set; }
    public string Path { get; set; } = "";
    public bool IsLocal { get; init; }
}

internal sealed class ClientHub
{
    private readonly object _gate = new();
    private readonly Dictionary<string, Viewer> _map = new(StringComparer.Ordinal);

    public void Touch(string? ip, string path)
    {
        var key = Normalize(ip);
        if (key.Length == 0)
        {
            return;
        }

        lock (_gate)
        {
            if (_map.TryGetValue(key, out var existing))
            {
                existing.LastSeen = DateTime.Now;
                existing.Path = path;
                return;
            }

            _map[key] = new Viewer
            {
                Ip = key,
                LastSeen = DateTime.Now,
                Path = path,
                IsLocal = key is "127.0.0.1" or "::1"
            };
        }
    }

    public List<Viewer> Active()
    {
        var cut = DateTime.Now.AddSeconds(-15);
        lock (_gate)
        {
            return _map.Values
                .Where(v => v.LastSeen >= cut)
                .OrderBy(v => v.IsLocal)
                .ThenByDescending(v => v.LastSeen)
                .ToList();
        }
    }

    private static string Normalize(string? ip)
    {
        if (string.IsNullOrWhiteSpace(ip))
        {
            return "";
        }

        if (ip.StartsWith("::ffff:", StringComparison.OrdinalIgnoreCase))
        {
            return ip[7..];
        }

        return ip == "::1" ? "127.0.0.1" : ip;
    }
}
