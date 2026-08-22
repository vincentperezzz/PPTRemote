namespace PptRemote;

internal static class AppIcon
{
    public static Icon Create()
    {
        var name = typeof(AppIcon).Assembly.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith("app.ico", StringComparison.OrdinalIgnoreCase));
        if (name != null)
        {
            using var stream = typeof(AppIcon).Assembly.GetManifestResourceStream(name);
            if (stream != null)
            {
                return new Icon(stream);
            }
        }

        var path = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(path) && File.Exists(path))
        {
            var associated = Icon.ExtractAssociatedIcon(path);
            if (associated != null)
            {
                return associated;
            }
        }

        return SystemIcons.Application;
    }
}
