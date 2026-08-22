using System.Collections.Concurrent;
using System.Runtime.InteropServices;

namespace PptRemote;

internal sealed class SlideInfo
{
    public int index { get; set; }
    public bool hidden { get; set; }
}

internal sealed class PresenterState
{
    public bool connected { get; set; }
    public bool slideshow { get; set; }
    public bool black { get; set; }
    public bool white { get; set; }
    public string title { get; set; } = "";
    public int index { get; set; }
    public int total { get; set; }
    public string notes { get; set; } = "";
    public int? nextIndex { get; set; }
    public bool thumbsReady { get; set; }
    public int thumbsVersion { get; set; }
    public List<SlideInfo> slides { get; set; } = new();
    public string message { get; set; } = "";
}

internal sealed class PowerPointService : IDisposable
{
    public const int PpPlaceholderBody = 2;
    public const int PpSlideShowRunning = 1;
    public const int PpSlideShowBlackScreen = 3;
    public const int PpSlideShowWhiteScreen = 4;
    private static readonly HashSet<int> SkipPlaceholders = new() { 1, 10, 11, 13, 15, 16 };

    public string ThumbDir { get; }
    private readonly object _gate = new();
    private readonly BlockingCollection<Action> _jobs = new();
    private readonly Thread _sta;
    private volatile bool _alive = true;
    private PresenterState _state;
    private dynamic? _app;
    private string _exportKey = "";
    private int _exportNext = 1;
    private int _notesNext = 1;
    private readonly Dictionary<int, string> _notes = new();
    private int _thumbsVersion;

    public PowerPointService()
    {
        ThumbDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PptRemote",
            "thumbs");
        Directory.CreateDirectory(ThumbDir);
        _state = Empty("Waiting for PowerPoint");
        _sta = new Thread(StaLoop)
        {
            Name = "ppt-com",
            IsBackground = true
        };
        _sta.SetApartmentState(ApartmentState.STA);
        _sta.Start();
    }

    public PresenterState Snapshot()
    {
        lock (_gate)
        {
            return Clone(_state);
        }
    }

    public PresenterState Next() => Run(() => ShowView(true).Next());
    public PresenterState Prev() => Run(() => ShowView(true).Previous());
    public PresenterState First() => Run(() => ShowView(true).First());
    public PresenterState Last() => Run(() => ShowView(true).Last());
    public PresenterState Goto(int index) => Run(() => ShowView(true).GotoSlide(index));

    public PresenterState StartShow() => Run(EnsureShow);

    public PresenterState EndShow() => Run(() =>
    {
        if (!TryApp())
        {
            return;
        }

        if ((int)_app!.SlideShowWindows.Count >= 1)
        {
            _app.SlideShowWindows.Item(1).View.Exit();
        }
    });

    public PresenterState Black() => Run(() => ToggleState(PpSlideShowBlackScreen));
    public PresenterState White() => Run(() => ToggleState(PpSlideShowWhiteScreen));

    public void Dispose()
    {
        _alive = false;
        try
        {
            _jobs.CompleteAdding();
        }
        catch
        {
        }

        Drop("");
    }

    private PresenterState Run(Action action)
    {
        var tcs = new TaskCompletionSource<PresenterState>(TaskCreationOptions.RunContinuationsAsynchronously);
        try
        {
            _jobs.Add(() =>
            {
                try
                {
                    action();
                    Poll();
                    tcs.TrySetResult(Snapshot());
                }
                catch (Exception ex)
                {
                    var snap = Snapshot();
                    snap.message = ex.Message;
                    tcs.TrySetResult(snap);
                }
            });
        }
        catch
        {
            return Snapshot();
        }

        return tcs.Task.Wait(8000) ? tcs.Task.Result : Snapshot();
    }

    private void StaLoop()
    {
        while (_alive)
        {
            Action? job = null;
            try
            {
                _jobs.TryTake(out job, 100);
            }
            catch
            {
            }

            if (job != null)
            {
                try
                {
                    job();
                }
                catch
                {
                }
            }

            try
            {
                Poll();
                if (job == null)
                {
                    ExportOne();
                    PrefetchNotes();
                }
            }
            catch
            {
                Drop("PowerPoint disconnected");
            }
        }
    }

    private void ToggleState(int target)
    {
        var view = ShowView(true);
        var current = (int)view.State;
        view.State = current == target ? PpSlideShowRunning : target;
    }

    private bool TryApp()
    {
        if (_app != null)
        {
            try
            {
                _ = _app.Presentations.Count;
                return true;
            }
            catch
            {
                ReleaseApp();
            }
        }

        try
        {
            _app = Com.GetActive("PowerPoint.Application");
            return true;
        }
        catch
        {
            return false;
        }
    }

    private void ReleaseApp()
    {
        if (_app == null)
        {
            return;
        }

        try
        {
            Marshal.ReleaseComObject(_app);
        }
        catch
        {
        }

        _app = null;
    }

    private void Drop(string message)
    {
        ReleaseApp();
        lock (_gate)
        {
            _state = Empty(message);
        }
    }

    private PresenterState Empty(string message) => new()
    {
        connected = false,
        message = message,
        thumbsVersion = _thumbsVersion
    };

    private void Poll()
    {
        if (!TryApp())
        {
            Drop("Waiting for PowerPoint");
            return;
        }

        if ((int)_app!.Presentations.Count < 1)
        {
            Drop("Open a presentation in PowerPoint");
            return;
        }

        var slideshow = (int)_app.SlideShowWindows.Count >= 1;
        dynamic pres;
        int index;
        var black = false;
        var white = false;
        if (slideshow)
        {
            var show = _app.SlideShowWindows.Item(1);
            pres = show.Presentation;
            var view = show.View;
            try
            {
                var st = (int)view.State;
                black = st == PpSlideShowBlackScreen;
                white = st == PpSlideShowWhiteScreen;
            }
            catch
            {
            }

            try
            {
                index = (int)view.Slide.SlideIndex;
            }
            catch
            {
                index = (int)view.CurrentShowPosition;
            }
        }
        else
        {
            pres = _app.ActivePresentation;
            try
            {
                index = (int)_app.ActiveWindow.View.Slide.SlideIndex;
            }
            catch
            {
                index = 1;
            }
        }

        var total = (int)pres.Slides.Count;
        if (index < 1)
        {
            index = 1;
        }

        if (total > 0 && index > total)
        {
            index = total;
        }

        var slides = new List<SlideInfo>(total);
        int? nextIndex = null;
        for (var i = 1; i <= total; i++)
        {
            var slide = pres.Slides.Item(i);
            var hidden = false;
            try
            {
                hidden = Flag(slide.SlideShowTransition.Hidden);
            }
            catch
            {
            }

            slides.Add(new SlideInfo { index = i, hidden = hidden });
            if (nextIndex == null && i > index && !hidden)
            {
                nextIndex = i;
            }
        }

        var notes = "";
        if (index >= 1 && index <= total)
        {
            notes = LookupNotes(pres, index);
        }

        string title;
        try
        {
            title = (string)pres.Name;
        }
        catch
        {
            title = "";
        }

        string full;
        try
        {
            full = (string)pres.FullName;
        }
        catch
        {
            full = title;
        }

        var key = full + "|" + total;
        if (key != _exportKey)
        {
            _exportKey = key;
            _exportNext = 1;
            _notesNext = 1;
            _notes.Clear();
            _thumbsVersion++;
            WipePngs();
        }

        var ready = total == 0 || _exportNext > total;
        lock (_gate)
        {
            _state = new PresenterState
            {
                connected = true,
                slideshow = slideshow,
                black = black,
                white = white,
                title = title,
                index = index,
                total = total,
                notes = notes,
                nextIndex = nextIndex,
                thumbsReady = ready,
                thumbsVersion = _thumbsVersion,
                slides = slides,
                message = slideshow ? "" : "Slideshow not running"
            };
        }
    }

    private string LookupNotes(dynamic pres, int index)
    {
        if (_notes.TryGetValue(index, out var cached))
        {
            return cached;
        }

        var text = ReadNotes(pres.Slides.Item(index));
        _notes[index] = text;
        return text;
    }

    private void PrefetchNotes()
    {
        if (!TryApp() || (int)_app!.Presentations.Count < 1)
        {
            return;
        }

        if ((int)_app.SlideShowWindows.Count >= 1)
        {
            return;
        }

        dynamic pres = _app.ActivePresentation;
        var total = (int)pres.Slides.Count;
        if (_notesNext > total)
        {
            return;
        }

        var i = _notesNext;
        try
        {
            _notes[i] = ReadNotes(pres.Slides.Item(i));
        }
        catch
        {
            _notes[i] = "";
        }

        _notesNext = i + 1;
    }

    private void ExportOne()
    {
        if (!TryApp() || (int)_app!.Presentations.Count < 1)
        {
            return;
        }

        if ((int)_app.SlideShowWindows.Count >= 1)
        {
            return;
        }

        dynamic pres = _app.ActivePresentation;
        var total = (int)pres.Slides.Count;
        if (_exportNext > total)
        {
            return;
        }

        var i = _exportNext;
        try
        {
            var slide = pres.Slides.Item(i);
            float w = (float)pres.PageSetup.SlideWidth;
            float h = (float)pres.PageSetup.SlideHeight;
            if (w <= 0)
            {
                w = 1;
            }

            var eh = Math.Max(1, (int)(960 * (h / w)));
            var dest = Path.Combine(ThumbDir, i + ".png");
            slide.Export(dest, "PNG", 960, eh);
        }
        catch
        {
        }

        _exportNext = i + 1;
        if (_exportNext > total)
        {
            lock (_gate)
            {
                _state.thumbsReady = true;
                _state.thumbsVersion = _thumbsVersion;
            }
        }
    }

    private void EnsureShow()
    {
        if (!TryApp())
        {
            throw new InvalidOperationException("PowerPoint is not open");
        }

        if ((int)_app!.SlideShowWindows.Count >= 1)
        {
            return;
        }

        if ((int)_app.Presentations.Count < 1)
        {
            throw new InvalidOperationException("Open a presentation first");
        }

        _app.ActivePresentation.SlideShowSettings.Run();
        Thread.Sleep(200);
    }

    private dynamic ShowView(bool create)
    {
        if ((int)_app!.SlideShowWindows.Count < 1)
        {
            if (!create)
            {
                throw new InvalidOperationException("Slideshow is not running");
            }

            EnsureShow();
        }

        return _app.SlideShowWindows.Item(1).View;
    }

    private static bool Flag(dynamic value)
    {
        try
        {
            return Convert.ToInt32(value) != 0;
        }
        catch
        {
            try
            {
                return (bool)value;
            }
            catch
            {
                return false;
            }
        }
    }

    private static string ReadNotes(dynamic slide)
    {
        try
        {
            try
            {
                var quick = ((string)slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text)
                    .Replace("\r", "\n")
                    .Replace("\v", "\n")
                    .Trim();
                if (quick.Length > 0)
                {
                    return quick;
                }
            }
            catch
            {
            }

            var shapes = slide.NotesPage.Shapes;
            var count = (int)shapes.Count;
            var body = "";
            var other = new List<string>();
            for (var i = 1; i <= count; i++)
            {
                var shape = shapes.Item(i);
                try
                {
                    if (!Flag(shape.HasTextFrame))
                    {
                        continue;
                    }

                    var text = ((string)shape.TextFrame.TextRange.Text).Replace("\r", "\n").Replace("\v", "\n").Trim();
                    if (text.Length == 0)
                    {
                        continue;
                    }

                    var isBody = false;
                    var ptype = -1;
                    try
                    {
                        ptype = (int)shape.PlaceholderFormat.Type;
                        isBody = ptype == PpPlaceholderBody;
                    }
                    catch
                    {
                    }

                    if (isBody)
                    {
                        body = text;
                    }
                    else if (!SkipPlaceholders.Contains(ptype))
                    {
                        other.Add(text);
                    }
                }
                catch
                {
                }
            }

            if (body.Length > 0)
            {
                return body;
            }

            if (other.Count == 0)
            {
                return "";
            }

            return other.OrderByDescending(s => s.Length).First();
        }
        catch
        {
            return "";
        }
    }

    private void WipePngs()
    {
        Directory.CreateDirectory(ThumbDir);
        foreach (var file in Directory.GetFiles(ThumbDir, "*.png"))
        {
            try
            {
                File.Delete(file);
            }
            catch
            {
            }
        }
    }

    private static PresenterState Clone(PresenterState s) => new()
    {
        connected = s.connected,
        slideshow = s.slideshow,
        black = s.black,
        white = s.white,
        title = s.title,
        index = s.index,
        total = s.total,
        notes = s.notes,
        nextIndex = s.nextIndex,
        thumbsReady = s.thumbsReady,
        thumbsVersion = s.thumbsVersion,
        slides = s.slides.ToList(),
        message = s.message
    };
}
