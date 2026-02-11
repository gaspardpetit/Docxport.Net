using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Frames;
using DocxportNet.Fields.Resolution;
using DocxportNet.Middleware;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Globalization;

namespace DocxportNet.Walker;

public sealed partial class DxpFieldEvalMiddleware : DxpLoggingMiddleware
{
	private DxpIVisitor _next { get; }
	public override DxpIVisitor Next => _currentAdapter ?? _next;


    private readonly DxpFieldEval _eval;
    private readonly DxpFieldEvalContext _context;
    private readonly DxpEvalFieldMode _mode;
    private readonly bool _includeDocumentProperties;
    private readonly bool _includeCustomProperties;
    private readonly Func<DateTimeOffset>? _nowProvider;
    private readonly ILogger? _logger;
    private readonly DxpEvalFieldMiddlewareOptions _options;
    private bool _initialized;
    private int _paragraphOrder;
    private readonly Stack<DxpIFieldEvalFrame> _fieldFrames = new();
    private readonly Stack<DxpIVisitor> _frameAdapters = new();
    private DxpIFieldEvalFrame? _outerFrame;
    private DxpIVisitor? _currentAdapter;

    public DxpFieldEvalMiddleware(
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpEvalFieldMode mode = DxpEvalFieldMode.Evaluate,
        bool includeDocumentProperties = true,
        bool includeCustomProperties = true,
        Func<DateTimeOffset>? nowProvider = null,
        ILogger? logger = null,
        DxpEvalFieldMiddlewareOptions? options = null)
        : base(logger, "DxpFieldEvalMiddleware")
    {
        _next = next ?? throw new ArgumentNullException(nameof(next));
		_eval = eval ?? throw new ArgumentNullException(nameof(eval));
        _context = _eval.Context;
        _mode = mode;
        _includeDocumentProperties = includeDocumentProperties;
        _includeCustomProperties = includeCustomProperties;
        _nowProvider = nowProvider;
        _logger = logger;
        _options = options ?? new DxpEvalFieldMiddlewareOptions();
    }

    public override IDisposable VisitDocumentBegin(WordprocessingDocument doc, DxpIDocumentContext documentContext)
    {
        if (!_initialized)
        {
            _context.InitFromDocumentContext(documentContext, _includeDocumentProperties, _includeCustomProperties);
            if (_nowProvider != null)
                _context.SetNow(_nowProvider);
            _context.TableResolver ??= new DxpWalkerTableResolver(documentContext);
            _context.RefResolver ??= _options.RefResolver ?? new DxpRefIndexResolver();
            _initialized = true;
        }

        _paragraphOrder = 0;
        _context.FieldDepth = 0;
        _context.OuterFrame = null;
        return Next.VisitDocumentBegin(doc, documentContext);
    }

    private DxpIFieldEvalFrame? CurrentField => _fieldFrames.Count > 0 ? _fieldFrames.Peek() : null;

    private DxpIFieldEvalFrame CreateInitialFrame(
        bool isComplex,
        bool inResult,
        bool seenSeparate,
        string? instructionText = null)
    {
        var next = GetChainedNext();
        var frame = new DxpEvalGenericFieldFrame(
            next,
            _eval,
            _context,
            _logger,
            _mode,
            inResult,
            seenSeparate,
            instructionText);
        return frame;
    }

    private void PushAdapterForFrame(DxpIFieldEvalFrame frame)
    {
        if (frame is DxpIVisitor visitor)
        {
            _frameAdapters.Push(visitor);
            _currentAdapter = visitor;
            return;
        }
        throw new InvalidOperationException($"Field frame '{frame.GetType().Name}' does not implement {nameof(DxpIVisitor)}.");
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
        // Document context state is maintained downstream.
    }

    private void UpdateFrameState()
    {
        _context.FieldDepth = _fieldFrames.Count;
        _context.OuterFrame = _outerFrame;
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("FrameState: depth={Depth} outer={Outer}",
                _context.FieldDepth,
                _context.OuterFrame?.GetType().Name ?? "null");
    }

    private DxpIVisitor GetChainedNext()
    {
        return _currentAdapter ?? _next;
    }

    public override void VisitComplexFieldBegin(FieldChar begin, DxpIDocumentContext d)
    {
        var frame = CreateInitialFrame(isComplex: true, inResult: false, seenSeparate: false);
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
    {
        _currentAdapter?.VisitComplexFieldInstruction(instr, text, d);
    }

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
        var frame = CreateInitialFrame(isComplex: false, inResult: true, seenSeparate: true, instruction);
        _fieldFrames.Push(frame);
        if (_fieldFrames.Count == 1)
            _outerFrame = frame;
        PushAdapterForFrame(frame);
        UpdateFrameState();
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("SimpleFieldBegin: frame={Frame} instruction='{Instruction}' depth={Depth}",
                frame.GetType().Name,
                instruction ?? string.Empty,
                _context.FieldDepth);

        var inner = _currentAdapter!.VisitSimpleFieldBegin(fld, d);
        return new DxpCompositeScope(inner, () => {
            PopCurrentField(d);
        });
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        LogTextWithFont("Eval.VisitText", t.Text);
        Next.VisitText(t, d);
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

    private void LogTextWithFont(string source, string text)
    {
        var run = _context.CurrentRun;
        string? fontSizeHp = run?.RunProperties?.FontSize?.Val?.Value;
        if (string.IsNullOrWhiteSpace(fontSizeHp))
            fontSizeHp = run?.Ancestors<Paragraph>()
                .FirstOrDefault()?
                .ParagraphProperties?
                .GetFirstChild<RunProperties>()?
                .FontSize?.Val?.Value;

        var fontSizePt = int.TryParse(fontSizeHp, NumberStyles.Integer, CultureInfo.InvariantCulture, out var hp)
            ? (hp / 2.0).ToString("0.###", CultureInfo.InvariantCulture)
            : "null";
        var escapedText = text
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t");
        Console.WriteLine($"[{source}] Text='{escapedText}' FontSizeHp={fontSizeHp ?? "null"} FontSizePt={fontSizePt}");
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        var previous = _context.Culture;
        var previousOutlineProvider = _context.CurrentOutlineLevelProvider;
        var previousOrder = _context.CurrentDocumentOrder;
        if (TryResolveParagraphCulture(p, d, _logger, out var culture))
            _context.Culture = culture;
        _context.CurrentOutlineLevelProvider = CreateOutlineLevelProvider(p, d);
        _context.CurrentDocumentOrder = ++_paragraphOrder;

        var inner = Next.VisitParagraphBegin(p, d, paragraph);
        return new DxpCompositeScope(inner, () => {
            _context.Culture = previous;
            _context.CurrentOutlineLevelProvider = previousOutlineProvider;
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

        if (CurrentField is DxpEvalGenericFieldFrame generic)
            generic.TryCaptureCodeRun(r);

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

        // Word stores outline levels as 0-based; SEQ \s expects 1-based.
        int resolved = level.HasValue ? level.Value + 1 : 0;
        return () => resolved;
    }

    private sealed class DxpCompositeScope : IDisposable
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
