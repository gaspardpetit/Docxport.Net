using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Frames;
using DocxportNet.Middleware;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace DocxportNet.Walker;

public abstract class DxpFieldMiddlewareBase : DxpLoggingMiddleware
{
    private readonly DxpIVisitor _next;
    public override DxpIVisitor Next => _currentAdapter ?? _next;

    private readonly DxpFieldEval _eval;
    private readonly DxpFieldEvalContext _context;
    private readonly bool _includeDocumentProperties;
    private readonly bool _includeCustomProperties;
    private readonly Func<DateTimeOffset>? _nowProvider;
    private readonly ILogger? _logger;
    private bool _initialized;
    private int _paragraphOrder;
    private readonly Stack<DxpIFieldEvalFrame> _fieldFrames = new();
    private readonly Stack<DxpIVisitor> _frameAdapters = new();
    private DxpIFieldEvalFrame? _outerFrame;
    private DxpIVisitor? _currentAdapter;

    protected DxpFieldMiddlewareBase(
        DxpIVisitor next,
        DxpFieldEval eval,
        bool includeDocumentProperties,
        bool includeCustomProperties,
        Func<DateTimeOffset>? nowProvider,
        ILogger? logger,
        string middlewareName)
        : base(logger, middlewareName)
    {
        _next = next ?? throw new ArgumentNullException(nameof(next));
        _eval = eval ?? throw new ArgumentNullException(nameof(eval));
        _context = _eval.Context;
        _includeDocumentProperties = includeDocumentProperties;
        _includeCustomProperties = includeCustomProperties;
        _nowProvider = nowProvider;
        _logger = logger;
    }

    protected DxpFieldEval Eval => _eval;
    protected DxpFieldEvalContext Context => _context;
    protected ILogger? Logger => _logger;

    internal abstract DxpIFieldEvalFrame CreateComplexFieldFrame();
    internal abstract DxpIFieldEvalFrame CreateSimpleFieldFrame(string? instructionText);
    protected virtual void InitializeModeSpecificContext(DxpIDocumentContext documentContext)
    {
    }

    protected DxpIVisitor GetChainedNext()
        => _currentAdapter ?? _next;

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        if (!_initialized)
        {
            _context.InitFromDocumentContext(documentContext, _includeDocumentProperties, _includeCustomProperties);
            if (_nowProvider != null)
                _context.SetNow(_nowProvider);
            _context.TableResolver ??= new DxpWalkerTableValueResolver(documentContext);
            InitializeModeSpecificContext(documentContext);
            _initialized = true;
        }

        _paragraphOrder = 0;
        _context.AutoNumbers.Reset();
        _context.FieldDepth = 0;
        _context.OuterFrame = null;
        return Next.VisitDocumentBegin(doc, documentContext);
    }

    private DxpIFieldEvalFrame? CurrentField => _fieldFrames.Count > 0 ? _fieldFrames.Peek() : null;

    private void PushAdapterForFrame(DxpIFieldEvalFrame frame)
    {
        if (frame is not DxpIVisitor visitor)
            throw new InvalidOperationException($"Field frame '{frame.GetType().Name}' does not implement {nameof(DxpIVisitor)}.");

        _frameAdapters.Push(visitor);
        _currentAdapter = visitor;
    }

    private void PopCurrentAdapter()
    {
        if (_frameAdapters.Count > 0)
            _frameAdapters.Pop();
        _currentAdapter = _frameAdapters.Count > 0 ? _frameAdapters.Peek() : null;
    }

    private void PopCurrentField(DxpIDocumentContext d)
    {
        if (_fieldFrames.Count == 1)
            _outerFrame = null;
        _fieldFrames.Pop();
        PopCurrentAdapter();
        UpdateFrameState();
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("FieldEnd: depth={Depth}", _context.FieldDepth);
    }

    private void UpdateFrameState()
    {
        _context.FieldDepth = _fieldFrames.Count;
        _context.OuterFrame = _outerFrame;
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug(
                "FrameState: depth={Depth} outer={Outer}",
                _context.FieldDepth,
                _context.OuterFrame?.GetType().Name ?? "null");
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
        var frame = CreateComplexFieldFrame();
        _fieldFrames.Push(frame);
        if (_fieldFrames.Count == 1)
            _outerFrame = frame;
        PushAdapterForFrame(frame);
        UpdateFrameState();
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("FieldBegin: frame={Frame} depth={Depth}", frame.GetType().Name, _context.FieldDepth);
        _currentAdapter!.VisitComplexFieldBegin(begin, d);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        LogTextWithFont("Eval.CachedResult", text);
        _currentAdapter?.VisitComplexFieldCachedResultText(text, d);
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
        => _currentAdapter?.VisitComplexFieldInstruction(instr, text, d);

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
        => _currentAdapter?.VisitComplexFieldSeparate(separate, d);

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        _currentAdapter?.VisitComplexFieldEnd(end, d);
        PopCurrentField(d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        var instruction = fld.Instruction?.Value;
        var frame = CreateSimpleFieldFrame(instruction);
        _fieldFrames.Push(frame);
        if (_fieldFrames.Count == 1)
            _outerFrame = frame;
        PushAdapterForFrame(frame);
        UpdateFrameState();
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug(
                "SimpleFieldBegin: frame={Frame} instruction='{Instruction}' depth={Depth}",
                frame.GetType().Name,
                instruction ?? string.Empty,
                _context.FieldDepth);

        var inner = _currentAdapter!.VisitSimpleFieldBegin(fld, d);
        return new DxpCompositeScope(inner, () => PopCurrentField(d));
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        LogTextWithFont("Eval.VisitText", t.Text);
        Next.VisitText(t, d);
    }

    public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
    {
        var target = _currentAdapter ?? Next;
        return target.VisitBlockBegin(child, d);
    }

    internal static bool NeedsPreserveSpace(string text)
    {
        if (text.Length == 0)
            return false;
        if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
            return true;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '\t' || ch == '\r' || ch == '\n')
                return true;
            if (ch == ' ' && i + 1 < text.Length && text[i + 1] == ' ')
                return true;
        }
        return false;
    }

    protected void LogTextWithFont(string source, string text)
    {
        if (_logger?.IsEnabled(LogLevel.Debug) != true)
            return;

        var run = _context.CurrentRun;
        string? fontSizeHp = run?.RunProperties?.FontSize?.Val?.Value;
        if (string.IsNullOrWhiteSpace(fontSizeHp))
        {
            fontSizeHp = run?.Ancestors<Paragraph>()
                .FirstOrDefault()?
                .ParagraphProperties?
                .GetFirstChild<RunProperties>()?
                .FontSize?.Val?.Value;
        }

        var fontSizePt = int.TryParse(fontSizeHp, NumberStyles.Integer, CultureInfo.InvariantCulture, out var hp)
            ? (hp / 2.0).ToString("0.###", CultureInfo.InvariantCulture)
            : "null";
        var escapedText = text
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t");
        _logger.LogDebug(
            "[{Source}] Text='{Text}' FontSizeHp={FontSizeHp} FontSizePt={FontSizePt}",
            source,
            escapedText,
            fontSizeHp ?? "null",
            fontSizePt);
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        var previous = _context.Culture;
        var previousOutlineProvider = _context.CurrentOutlineLevelProvider;
        var previousHeadingProvider = _context.CurrentBuiltInHeadingLevelProvider;
        var previousStoryProvider = _context.CurrentStoryKeyProvider;
        var previousOrder = _context.CurrentDocumentOrder;
        if (TryResolveParagraphCulture(p, d, _logger, out var culture))
            _context.Culture = culture;
        _context.CurrentOutlineLevelProvider = CreateOutlineLevelProvider(p, d);
        _context.CurrentBuiltInHeadingLevelProvider = CreateBuiltInHeadingLevelProvider(p, d);
        _context.CurrentStoryKeyProvider = CreateStoryKeyProvider(p, d);
        _context.CurrentDocumentOrder = ++_paragraphOrder;

        var target = _currentAdapter ?? Next;
        var inner = target.VisitParagraphBegin(p, d, paragraph);
        return new DxpCompositeScope(inner, () => {
            _context.Culture = previous;
            _context.CurrentOutlineLevelProvider = previousOutlineProvider;
            _context.CurrentBuiltInHeadingLevelProvider = previousHeadingProvider;
            _context.CurrentStoryKeyProvider = previousStoryProvider;
            _context.CurrentDocumentOrder = previousOrder;
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        var previous = _context.Culture;
        var previousRun = _context.CurrentRun;
        _context.CurrentRun = r;
        LogTextWithFont("VisitRunBegin", r.InnerText);

        if (TryResolveRunCulture(r, d, _logger, out var culture))
            _context.Culture = culture;

        if (CurrentField is IDxpFieldCodeRunCapture capture)
            capture.TryCaptureCodeRun(r);

        var inner = Next.VisitRunBegin(r, d);
        return new DxpCompositeScope(inner, () => {
            _context.Culture = previous;
            _context.CurrentRun = previousRun;
        });
    }

    private static bool TryResolveParagraphCulture(Paragraph p, DxpIDocumentContext d, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        string? lang = null;

        if (d.Styles is DxpStyleResolver resolver)
            lang = resolver.ResolveParagraphLanguage(p) ?? resolver.GetDefaultLanguage();
        else
            lang = p.ParagraphProperties?.GetFirstChild<ParagraphMarkRunProperties>()
                ?.GetFirstChild<Languages>()?.Val?.Value;

        return TryCreateCulture(lang, logger, out culture);
    }

    private bool TryResolveRunCulture(Run r, DxpIDocumentContext d, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        string? lang = null;

        if (d.Styles is DxpStyleResolver resolver)
        {
            var paragraph = r.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph != null)
                lang = resolver.ResolveRunLanguage(paragraph, r);
        }

        lang ??= _context.CurrentRun?.RunProperties?.Languages?.Val?.Value ?? r.RunProperties?.Languages?.Val?.Value;
        return TryCreateCulture(lang, logger, out culture);
    }

    private static bool TryCreateCulture(string? lang, ILogger? logger, out CultureInfo culture)
    {
        culture = CultureInfo.CurrentCulture;
        if (string.IsNullOrWhiteSpace(lang))
            return false;

        try
        {
            culture = new CultureInfo(lang);
            return true;
        }
        catch (CultureNotFoundException)
        {
            logger?.LogWarning("Invalid language tag '{Lang}' in document; using current culture.", lang);
            return false;
        }
    }

    private static Func<int> CreateOutlineLevelProvider(Paragraph p, DxpIDocumentContext d)
    {
        int? level = null;
        if (d.Styles is DxpStyleResolver resolver)
            level = resolver.GetOutlineLevel(p);
        else
            level = p.ParagraphProperties?.OutlineLevel?.Val?.Value;

        int resolved = level.HasValue ? level.Value + 1 : 0;
        return () => resolved;
    }

    private static Func<int> CreateBuiltInHeadingLevelProvider(Paragraph p, DxpIDocumentContext d)
    {
        int level = 0;
        foreach (var style in d.Styles.GetParagraphStyleChain(p))
        {
            var id = style.StyleId;
            if (id.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(id.Substring("Heading".Length), out var parsed)
                && parsed is >= 1 and <= 9)
            {
                level = parsed;
                break;
            }
        }
        return () => level;
    }

    private static Func<string> CreateStoryKeyProvider(Paragraph p, DxpIDocumentContext d)
    {
        string part = d.CurrentPart?.Uri.ToString() ?? "main";
        var textBox = p.Ancestors<TextBoxContent>().FirstOrDefault();
        if (textBox != null)
            part += "#textbox:" + RuntimeHelpers.GetHashCode(textBox).ToString(CultureInfo.InvariantCulture);
        else if (p.Ancestors<Footnote>().FirstOrDefault()?.Id?.Value is long footnoteId)
            part += "#footnote:" + footnoteId.ToString(CultureInfo.InvariantCulture);
        else if (p.Ancestors<Endnote>().FirstOrDefault()?.Id?.Value is long endnoteId)
            part += "#endnote:" + endnoteId.ToString(CultureInfo.InvariantCulture);
        return () => part;
    }

    protected sealed class DxpCompositeScope : IDisposable
    {
        private readonly IDisposable _inner;
        private readonly Action _onDispose;
        private bool _disposed;

        public DxpCompositeScope(IDisposable inner, Action onDispose)
        {
            _inner = inner;
            _onDispose = onDispose;
        }

        public void Dispose()
        {
            if (_disposed)
                return;
            _disposed = true;
            _onDispose();
            _inner.Dispose();
        }
    }
}
