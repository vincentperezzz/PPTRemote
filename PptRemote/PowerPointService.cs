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
    public int clickIndex { get; set; }
    public int clickCount { get; set; }
    public int aheadVersion { get; set; }
    public bool aheadReady { get; set; }
    public string aheadLabel { get; set; } = "";
    public List<SlideInfo> slides { get; set; } = new();
    public string message { get; set; } = "";
}

internal sealed class PowerPointService : IDisposable
{
    public const int PpPlaceholderBody = 2;
    public const int PpSlideShowRunning = 1;
    public const int PpSlideShowBlackScreen = 3;
    public const int PpSlideShowWhiteScreen = 4;
    public const int MsoAnimTriggerOnPageClick = 1;
    private static readonly HashSet<int> SkipPlaceholders = new() { 1, 10, 11, 13, 15, 16 };

    public string ThumbDir { get; }
    public string AheadFile => Path.Combine(ThumbDir, "ahead.png");
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
    private readonly Dictionary<int, string> _slideSig = new();
    private readonly string _metaFile;
    private int _thumbsVersion;
    private int _aheadVersion;
    private string _aheadKey = "";
    private bool _aheadReady;
    private string _aheadLabel = "";
    private int _clickIndex;
    private int _clickCount;
    private int _aheadSlide;
    private int _aheadClicks;

    public PowerPointService()
    {
        ThumbDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PptRemote",
            "thumbs");
        Directory.CreateDirectory(ThumbDir);
        _metaFile = Path.Combine(ThumbDir, "deck.key");
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
                    ExportAhead();
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
        var clickIndex = 0;
        var clickCount = 0;
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

            clickIndex = ReadClick(view, true);
            clickCount = ReadClick(view, false);
        }
        else
        {
            pres = BestPresentation();
            try
            {
                index = (int)_app.ActiveWindow.View.Slide.SlideIndex;
            }
            catch
            {
                index = 1;
            }

            try
            {
                if ((int)pres.Slides.Count != (int)_app.ActivePresentation.Slides.Count)
                {
                    index = 1;
                }
            }
            catch
            {
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

        var key = DeckKey(full, total);
        if (key != _exportKey)
        {
            _exportKey = key;
            _notesNext = 1;
            _notes.Clear();
            _slideSig.Clear();
            _thumbsVersion++;
            _aheadVersion++;
            _aheadKey = "";
            _aheadReady = false;
            var disk = ReadMeta();
            if (disk != key)
            {
                _exportNext = 1;
                WipePngs();
                WriteMeta(key);
            }
            else
            {
                _exportNext = 1;
                while (_exportNext <= total && File.Exists(Path.Combine(ThumbDir, _exportNext + ".png")))
                {
                    _exportNext++;
                }
            }
        }

        WatchSlide(pres, index);
        if (nextIndex != null)
        {
            WatchSlide(pres, nextIndex.Value);
        }

        var aheadSlide = index;
        var aheadClicks = 0;
        var aheadLabel = "";
        if (slideshow && clickIndex < clickCount)
        {
            aheadSlide = index;
            aheadClicks = clickIndex + 1;
            aheadLabel = "Next click · " + aheadClicks + " / " + clickCount;
        }
        else if (nextIndex != null)
        {
            aheadSlide = nextIndex.Value;
            aheadClicks = 0;
            aheadLabel = "Next slide · " + aheadSlide + " / " + total;
        }
        else
        {
            aheadSlide = index;
            aheadClicks = clickCount;
            aheadLabel = "Last click";
        }

        _clickIndex = clickIndex;
        _clickCount = clickCount;
        _aheadSlide = aheadSlide;
        _aheadClicks = aheadClicks;
        _aheadLabel = aheadLabel;

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
                clickIndex = clickIndex,
                clickCount = clickCount,
                aheadVersion = _aheadVersion,
                aheadReady = _aheadReady,
                aheadLabel = aheadLabel,
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

        dynamic pres = BestPresentation();
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

        dynamic pres = (int)_app.SlideShowWindows.Count >= 1
            ? _app.SlideShowWindows.Item(1).Presentation
            : BestPresentation();
        var total = (int)pres.Slides.Count;
        if (_exportNext > total)
        {
            return;
        }

        var i = _exportNext;
        ExportSlide(pres, i);
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

    private dynamic BestPresentation()
    {
        dynamic? best = null;
        var bestCount = -1;
        var n = (int)_app!.Presentations.Count;
        for (var i = 1; i <= n; i++)
        {
            var p = _app.Presentations.Item(i);
            var c = 0;
            try
            {
                c = (int)p.Slides.Count;
            }
            catch
            {
            }

            if (c > bestCount)
            {
                best = p;
                bestCount = c;
            }
        }

        if (best == null)
        {
            throw new InvalidOperationException("Open a presentation first");
        }

        return best;
    }

    private void EnsureShow()
    {
        if (!TryApp())
        {
            throw new InvalidOperationException("PowerPoint is not open");
        }

        if ((int)_app!.Presentations.Count < 1)
        {
            throw new InvalidOperationException("Open a presentation first");
        }

        var best = BestPresentation();
        if ((int)_app.SlideShowWindows.Count >= 1)
        {
            dynamic showing = _app.SlideShowWindows.Item(1).Presentation;
            var showingCount = 0;
            try
            {
                showingCount = (int)showing.Slides.Count;
            }
            catch
            {
            }

            if (showingCount >= (int)best.Slides.Count)
            {
                return;
            }

            _app.SlideShowWindows.Item(1).View.Exit();
            Thread.Sleep(200);
        }

        best.SlideShowSettings.Run();
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
                var quick = CleanNotes((string)slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text);
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

                    var text = CleanNotes((string)shape.TextFrame.TextRange.Text);
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

    private static string CleanNotes(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return "";
        }

        var chars = text.Replace("\r\n", "\n").Replace('\r', '\n').ToCharArray();
        for (var i = 0; i < chars.Length; i++)
        {
            var c = chars[i];
            if (c != '\n' && c != '\t' && c < 32)
            {
                chars[i] = '\n';
            }
        }

        return new string(chars).Trim();
    }

    private static int ReadClick(dynamic view, bool index)
    {
        try
        {
            var n = index ? (int)view.GetClickIndex() : (int)view.GetClickCount();
            return n < 0 ? 0 : n;
        }
        catch
        {
            try
            {
                var n = index ? (int)view.GetClickIndex : (int)view.GetClickCount;
                return n < 0 ? 0 : n;
            }
            catch
            {
                return 0;
            }
        }
    }

    private void ExportAhead()
    {
        if (!TryApp() || (int)_app!.Presentations.Count < 1 || _aheadSlide < 1)
        {
            return;
        }

        var key = _exportKey + "|" + _aheadSlide + "|" + _aheadClicks;
        if (key == _aheadKey && File.Exists(AheadFile))
        {
            return;
        }

        dynamic pres;
        try
        {
            pres = (int)_app.SlideShowWindows.Count >= 1
                ? _app.SlideShowWindows.Item(1).Presentation
                : BestPresentation();
        }
        catch
        {
            return;
        }

        try
        {
            RenderAhead(pres, _aheadSlide, _aheadClicks);
            _aheadKey = key;
            _aheadVersion++;
            _aheadReady = true;
            lock (_gate)
            {
                _state.aheadVersion = _aheadVersion;
                _state.aheadReady = true;
                _state.aheadLabel = _aheadLabel;
            }
        }
        catch
        {
            try
            {
                var fallback = Path.Combine(ThumbDir, _aheadSlide + ".png");
                if (File.Exists(fallback))
                {
                    File.Copy(fallback, AheadFile, true);
                    _aheadKey = key;
                    _aheadVersion++;
                    _aheadReady = true;
                    lock (_gate)
                    {
                        _state.aheadVersion = _aheadVersion;
                        _state.aheadReady = true;
                        _state.aheadLabel = _aheadLabel;
                    }
                }
            }
            catch
            {
            }
        }
    }

    private void RenderAhead(dynamic pres, int slideIndex, int clicksToApply)
    {
        dynamic? temp = null;
        try
        {
            dynamic src = pres.Slides.Item(slideIndex);
            temp = _app!.Presentations.Add(0);
            src.Copy();
            temp.Slides.Paste();
            dynamic copy = temp.Slides.Item((int)temp.Slides.Count);
            ApplyClickVisibility(copy, clicksToApply);
            float w = (float)temp.PageSetup.SlideWidth;
            float h = (float)temp.PageSetup.SlideHeight;
            if (w <= 0)
            {
                w = 1;
            }

            var eh = Math.Max(1, (int)(960 * (h / w)));
            copy.Export(AheadFile, "PNG", 960, eh);
            temp.Saved = -1;
            temp.Close();
            temp = null;
        }
        finally
        {
            if (temp != null)
            {
                try
                {
                    temp.Saved = -1;
                    temp.Close();
                }
                catch
                {
                }
            }
        }
    }

    private static void ApplyClickVisibility(dynamic slide, int clicksToApply)
    {
        var entrance = new HashSet<int>();
        var effects = new List<(int trigger, bool exit, int id)>();
        try
        {
            var sequence = slide.TimeLine.MainSequence;
            var n = (int)sequence.Count;
            for (var i = 1; i <= n; i++)
            {
                var e = sequence.Item(i);
                var trigger = 0;
                try
                {
                    trigger = (int)e.Timing.TriggerType;
                }
                catch
                {
                }

                var exit = false;
                try
                {
                    exit = Flag(e.Exit);
                }
                catch
                {
                }

                int id;
                try
                {
                    id = (int)e.Shape.Id;
                }
                catch
                {
                    continue;
                }

                if (!exit)
                {
                    entrance.Add(id);
                }

                effects.Add((trigger, exit, id));
            }
        }
        catch
        {
            return;
        }

        var vis = new Dictionary<int, bool>();
        var shapes = (int)slide.Shapes.Count;
        for (var i = 1; i <= shapes; i++)
        {
            var sh = slide.Shapes.Item(i);
            int id;
            try
            {
                id = (int)sh.Id;
            }
            catch
            {
                continue;
            }

            vis[id] = !entrance.Contains(id);
        }

        var click = 0;
        foreach (var e in effects)
        {
            if (e.trigger == MsoAnimTriggerOnPageClick)
            {
                click++;
            }

            if (click > clicksToApply)
            {
                break;
            }

            vis[e.id] = !e.exit;
        }

        for (var i = 1; i <= shapes; i++)
        {
            var sh = slide.Shapes.Item(i);
            try
            {
                var id = (int)sh.Id;
                sh.Visible = vis.TryGetValue(id, out var on) && on ? -1 : 0;
            }
            catch
            {
            }
        }
    }

    private static string DeckKey(string full, int total)
    {
        var ticks = 0L;
        try
        {
            if (full.Length > 2 && File.Exists(full))
            {
                ticks = File.GetLastWriteTimeUtc(full).Ticks;
            }
        }
        catch
        {
        }

        return full + "|" + total + "|" + ticks;
    }

    private string ReadMeta()
    {
        try
        {
            return File.Exists(_metaFile) ? File.ReadAllText(_metaFile) : "";
        }
        catch
        {
            return "";
        }
    }

    private void WriteMeta(string key)
    {
        try
        {
            File.WriteAllText(_metaFile, key);
        }
        catch
        {
        }
    }

    private void WatchSlide(dynamic pres, int index)
    {
        if (index < 1)
        {
            return;
        }

        string sig;
        try
        {
            sig = SlideSig(pres.Slides.Item(index));
        }
        catch
        {
            return;
        }

        if (_slideSig.TryGetValue(index, out var old) && old == sig)
        {
            return;
        }

        var seen = _slideSig.ContainsKey(index);
        _slideSig[index] = sig;
        if (!seen)
        {
            return;
        }

        ExportSlide(pres, index);
        _thumbsVersion++;
        _aheadKey = "";
        _aheadReady = false;
        lock (_gate)
        {
            _state.thumbsVersion = _thumbsVersion;
        }
    }

    private static string SlideSig(dynamic slide)
    {
        var n = (int)slide.Shapes.Count;
        var acc = unchecked(n * 397);
        for (var i = 1; i <= n; i++)
        {
            try
            {
                var sh = slide.Shapes.Item(i);
                acc = unchecked(acc * 31 + (int)sh.Id);
                acc = unchecked(acc * 31 + (int)(float)sh.Width);
                acc = unchecked(acc * 31 + (int)(float)sh.Height);
                acc = unchecked(acc * 31 + (int)(float)sh.Left);
                acc = unchecked(acc * 31 + (int)(float)sh.Top);
            }
            catch
            {
            }
        }

        return acc.ToString();
    }

    private void ExportSlide(dynamic pres, int index)
    {
        try
        {
            var slide = pres.Slides.Item(index);
            float w = (float)pres.PageSetup.SlideWidth;
            float h = (float)pres.PageSetup.SlideHeight;
            if (w <= 0)
            {
                w = 1;
            }

            var eh = Math.Max(1, (int)(960 * (h / w)));
            var dest = Path.Combine(ThumbDir, index + ".png");
            var tmp = dest + ".tmp";
            try
            {
                File.Delete(tmp);
            }
            catch
            {
            }

            slide.Export(tmp, "PNG", 960, eh);
            try
            {
                File.Delete(dest);
            }
            catch
            {
            }

            File.Move(tmp, dest, true);
            try
            {
                _slideSig[index] = SlideSig(slide);
            }
            catch
            {
            }
        }
        catch
        {
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
        clickIndex = s.clickIndex,
        clickCount = s.clickCount,
        aheadVersion = s.aheadVersion,
        aheadReady = s.aheadReady,
        aheadLabel = s.aheadLabel,
        slides = s.slides.ToList(),
        message = s.message
    };
}
