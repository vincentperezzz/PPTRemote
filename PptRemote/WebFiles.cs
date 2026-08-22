using System.Reflection;

namespace PptRemote;

internal static class WebFiles
{
    private static readonly Assembly Asm = typeof(WebFiles).Assembly;
    private static readonly Dictionary<string, string> TextCache = new(StringComparer.OrdinalIgnoreCase);

    public static string Text(string file)
    {
        lock (TextCache)
        {
            if (TextCache.TryGetValue(file, out var cached))
            {
                return cached;
            }

            using var stream = Open(file);
            using var reader = new StreamReader(stream);
            var text = reader.ReadToEnd();
            TextCache[file] = text;
            return text;
        }
    }

    private static Stream Open(string file)
    {
        var match = Asm.GetManifestResourceNames()
            .FirstOrDefault(n => n.EndsWith("." + file, StringComparison.OrdinalIgnoreCase));
        if (match == null)
        {
            throw new FileNotFoundException("Missing web file " + file);
        }

        return Asm.GetManifestResourceStream(match)
            ?? throw new FileNotFoundException("Missing web file " + file);
    }
}
