using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

var dest = args.Length > 0
    ? args[0]
    : Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "PptRemote", "app.ico"));

var sizes = new[] { 16, 20, 24, 32, 40, 48, 64, 128, 256 };
var pngs = new List<byte[]>();
foreach (var size in sizes)
{
    using var bmp = Draw(size);
    using var ms = new MemoryStream();
    bmp.Save(ms, ImageFormat.Png);
    pngs.Add(ms.ToArray());
}

Directory.CreateDirectory(Path.GetDirectoryName(dest)!);
using var fs = File.Create(dest);
using var bw = new BinaryWriter(fs);
bw.Write((ushort)0);
bw.Write((ushort)1);
bw.Write((ushort)pngs.Count);
var offset = 6 + 16 * pngs.Count;
for (var i = 0; i < pngs.Count; i++)
{
    var size = sizes[i];
    var dim = (byte)(size >= 256 ? 0 : size);
    bw.Write(dim);
    bw.Write(dim);
    bw.Write((byte)0);
    bw.Write((byte)0);
    bw.Write((ushort)1);
    bw.Write((ushort)32);
    bw.Write(pngs[i].Length);
    bw.Write(offset);
    offset += pngs[i].Length;
}

foreach (var png in pngs)
{
    bw.Write(png);
}

Console.WriteLine(dest + " " + new FileInfo(dest).Length);

static Bitmap Draw(int size)
{
    var bmp = new Bitmap(size, size, PixelFormat.Format32bppArgb);
    using var g = Graphics.FromImage(bmp);
    g.SmoothingMode = SmoothingMode.AntiAlias;
    g.CompositingQuality = CompositingQuality.HighQuality;
    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
    g.Clear(Color.Transparent);
    var pad = Math.Max(1, size * 0.06f);
    using var fill = new SolidBrush(Color.FromArgb(255, 77, 46));
    g.FillEllipse(fill, pad, pad, size - pad * 2, size - pad * 2);
    var left = (int)(size * 0.38);
    var top = (int)(size * 0.28);
    var bot = (int)(size * 0.72);
    var right = (int)(size * 0.74);
    var mid = (int)(size * 0.50);
    g.FillPolygon(Brushes.White, new[]
    {
        new Point(left, top),
        new Point(left, bot),
        new Point(right, mid)
    });
    return bmp;
}
