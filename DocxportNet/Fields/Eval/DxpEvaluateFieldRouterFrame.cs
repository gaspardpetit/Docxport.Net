using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Frames;
using DocxportNet.Fields.Semantic;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpEvaluateFieldRouterFrame : DxpMiddleware, DxpIFieldEvalFrame, IDxpFieldCodeRunCapture, IDxpStructuredFieldResultSink, IDxpNestedFieldResultSink, IDxpDeferredFieldSink
{
    private static readonly DxpEvaluateFieldFrameFactory FrameFactory = new();

    public bool Evaluated { get; set; }
    public string? InstructionText { get; set; }
    public Run? CodeRun { get; set; }

    public override DxpIVisitor Next { get; }
    public DxpFieldEvalContext EvalContext { get; }

    private readonly DxpFieldEval _eval;
    private readonly ILogger? _logger;
    private bool _seenSeparate;
    private bool _inResult;
    private readonly List<FieldEvent> _events = new();
    private readonly Stack<IDisposable> _replayScopes = new();
    private DxpFieldValue? _capturedSetScalar;
    private bool _capturedSetScalarIsExact;
    private int _nestedSetResultCount;
    private readonly bool _isSimpleField;
    private readonly bool _allowDeferredCapture;
    private readonly List<DxpFieldExpressionPart> _expressionParts = new();

    public DxpEvaluateFieldRouterFrame(
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext evalContext,
        ILogger? logger,
        bool initialInResult = false,
        bool initialSeenSeparate = false,
        string? initialInstructionText = null,
        bool allowDeferredCapture = true)
        : base()
    {
        Next = next ?? throw new ArgumentNullException(nameof(next));
        _eval = eval ?? throw new ArgumentNullException(nameof(eval));
        EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
        _inResult = initialInResult;
        _seenSeparate = initialSeenSeparate;
        InstructionText = initialInstructionText;
        _isSimpleField = initialInResult && initialSeenSeparate;
        _allowDeferredCapture = allowDeferredCapture;
        if (_isSimpleField && !string.IsNullOrEmpty(initialInstructionText))
            _expressionParts.Add(new DxpFieldExpressionText(initialInstructionText));
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || _inResult)
            return;
        _events.Add(FieldEvent.Instruction(instr, text));
        AppendInstructionText(text);
        _expressionParts.Add(new DxpFieldExpressionText(text));
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!_seenSeparate)
        {
            _seenSeparate = true;
            _inResult = true;
        }

        _events.Add(FieldEvent.Separate(separate));
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        _events.Add(FieldEvent.End(end));
        ReplayEvents(d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        _events.Add(FieldEvent.SimpleBegin(fld));
        return DxpDisposable.Create(() => {
            _events.Add(FieldEvent.SimpleEnd());
            ReplayEvents(d);
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (!_inResult)
            return DxpDisposable.Empty;
        _events.Add(FieldEvent.RunBegin(r));
        return DxpDisposable.Create(() => _events.Add(FieldEvent.RunEnd()));
    }

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        if (!_inResult)
            return DxpDisposable.Empty;
        _events.Add(FieldEvent.HyperlinkBegin(link, target));
        return DxpDisposable.Create(() => _events.Add(FieldEvent.HyperlinkEnd()));
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (!_inResult)
        {
            AppendInstructionText(t.Text);
            _events.Add(FieldEvent.Text(t));
            _expressionParts.Add(new DxpFieldExpressionText(t.Text));
            return;
        }
        _events.Add(FieldEvent.Text(t));
    }

    public override void VisitBreak(Break br, DxpIDocumentContext d)
    {
        if (!_inResult)
        {
            AppendInstructionText("\n");
            _events.Add(FieldEvent.Break(br));
            _expressionParts.Add(new DxpFieldExpressionText("\n"));
            return;
        }
        _events.Add(FieldEvent.Break(br));
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        if (!_inResult)
        {
            AppendInstructionText("\t");
            _events.Add(FieldEvent.Tab(tab));
            _expressionParts.Add(new DxpFieldExpressionText("\t"));
            return;
        }
        _events.Add(FieldEvent.Tab(tab));
    }

    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        if (!_inResult)
        {
            AppendInstructionText("\n");
            _events.Add(FieldEvent.CarriageReturn(cr));
            _expressionParts.Add(new DxpFieldExpressionText("\n"));
            return;
        }
        _events.Add(FieldEvent.CarriageReturn(cr));
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
    {
        if (!_inResult)
        {
            AppendInstructionText("-");
            _events.Add(FieldEvent.NoBreakHyphen(nbh));
            _expressionParts.Add(new DxpFieldExpressionText("-"));
            return;
        }
        _events.Add(FieldEvent.NoBreakHyphen(nbh));
    }

    public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
    {
        if (_inResult)
            return DxpDisposable.Empty;
        return Next.VisitBlockBegin(child, d);
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        if (!_inResult)
        {
            _events.Add(FieldEvent.ParagraphBegin(p, paragraph));
            if (_expressionParts.Count > 0)
                _expressionParts.Add(new DxpFieldExpressionParagraph());
            if (EvalContext.SuppressCrossParagraphFieldParagraphOutput)
                return DxpDisposable.Create(() => _events.Add(FieldEvent.ParagraphEnd()));
            var inner = Next.VisitParagraphBegin(p, d, paragraph);
            return DxpDisposable.Create(() => {
                _events.Add(FieldEvent.ParagraphEnd());
                inner.Dispose();
            });
        }

        _events.Add(FieldEvent.ParagraphBegin(p, paragraph));
        return DxpDisposable.Create(() => _events.Add(FieldEvent.ParagraphEnd()));
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
        => false;

    private void EmitUnsupported(DxpIDocumentContext d)
    {
        if (Evaluated)
            return;
        Evaluated = true;

        var instruction = string.IsNullOrWhiteSpace(InstructionText) ? " " : InstructionText!;
        var text = DxpFieldEvalRules.GetEvaluationErrorText(instruction);
        var t = new Text(text);

        var run = new Run();
        using (Next.VisitRunBegin(run, d))
            Next.VisitText(t, d);
    }

    private void AppendInstructionText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;
        InstructionText = InstructionText == null ? text : InstructionText + text;
    }

    private void ReplayEvents(DxpIDocumentContext d)
    {
        if (_allowDeferredCapture &&
            Next is IDxpDeferredFieldSink deferredSink &&
            !string.IsNullOrWhiteSpace(InstructionText) &&
            deferredSink.TryRecordDeferredField(CreateDeferredField(), d))
        {
            _events.Clear();
            return;
        }

        if (TryReplaySemanticExpression(d))
        {
            _events.Clear();
            return;
        }

        DxpIFieldEvalFrame? delegateFrame = FrameFactory.Create(InstructionText, Next, _eval, EvalContext, _logger);
        if (delegateFrame == null)
        {
            EmitUnsupported(d);
            _events.Clear();
            return;
        }

        if (delegateFrame is DxpSetFieldEvalFrame setFrame && _capturedSetScalarIsExact)
            setFrame.CapturedScalar = _capturedSetScalar;

        if (_logger?.IsEnabled(LogLevel.Debug) == true)
        {
            _logger.LogDebug(
                "Generic.Replay: frame={Frame} mode=Evaluate events={EventCount}",
                delegateFrame.GetType().Name,
                _events.Count);
        }

        foreach (var ev in _events)
            ev.Replay(delegateFrame, d, _replayScopes);
        while (_replayScopes.Count > 0)
            _replayScopes.Pop().Dispose();
        _events.Clear();
    }

    private bool TryReplaySemanticExpression(DxpIDocumentContext context)
    {
        if (!EvalContext.UseSemanticFieldResults)
            return false;
        if (EvalContext.PreserveLayoutDependentFields &&
            DxpFieldInstructionClassifier.IsPaginationDependentInstruction(InstructionText))
            return false;

        var evaluator = new DxpSemanticFieldEvaluator(_eval);
        DxpSemanticFieldResult result = evaluator.EvaluateExpressionAsync(
            new DxpFieldExpression(_expressionParts.ToArray()), context).GetAwaiter().GetResult();
        if (result.Status == DxpFieldEvalStatus.Failed)
            return false;

        Evaluated = true;
        DxpFieldNodeBuffer buffer = DxpSemanticFieldResultAdapter.BuildBuffer(
            result.Content, CodeRun, _eval, _logger);
        if (buffer.IsEmpty)
            return true;
        if (buffer.HasBlockRoots && EvalContext.FieldDepth == 1)
            EvalContext.DeferStructuredFieldResult(buffer);
        else
            buffer.Replay(Next, context);
        return true;
    }

    public void TryCaptureCodeRun(Run r)
    {
        if (CodeRun == null && !_inResult)
            CodeRun = DxpRunCloner.CloneRunWithParagraphAncestor(r);
    }

    public bool TryRecordStructuredFieldResult(DxpFieldNodeBuffer buffer)
    {
        _events.Add(FieldEvent.StructuredResult(buffer));
        return true;
    }

    public bool TryRecordDeferredField(DxpDeferredField field, DxpIDocumentContext context)
    {
        _ = context;
        bool parentIsIf = DxpFieldInstructionClassifier.IsIfInstruction(InstructionText);
        if (!parentIsIf && !EvalContext.UseSemanticFieldResults)
            return false;

        _events.Add(FieldEvent.DeferredField(field));
        _expressionParts.Add(new DxpFieldExpressionChild(field.Expression));
        return true;
    }

    private DxpDeferredField CreateDeferredField()
    {
        string instruction = InstructionText ?? string.Empty;
        FieldEvent[] events = _events.ToArray();
        bool isSimpleField = _isSimpleField;
        DxpFieldValue? capturedSetScalar = _capturedSetScalar;
        bool capturedSetScalarIsExact = _capturedSetScalarIsExact;
        int nestedSetResultCount = _nestedSetResultCount;
        return new DxpDeferredField(
            instruction,
            new DxpFieldExpression(_expressionParts.ToArray()),
            (visitor, context) => {
                var router = new DxpEvaluateFieldRouterFrame(
                    visitor,
                    _eval,
                    EvalContext,
                    _logger,
                    initialInResult: isSimpleField,
                    initialSeenSeparate: isSimpleField,
                    initialInstructionText: isSimpleField ? instruction : null,
                    allowDeferredCapture: false);
                router._capturedSetScalar = capturedSetScalar;
                router._capturedSetScalarIsExact = capturedSetScalarIsExact;
                router._nestedSetResultCount = nestedSetResultCount;
                var scopes = new Stack<IDisposable>();
                foreach (var fieldEvent in events)
                    fieldEvent.Replay(router, context, scopes);
                while (scopes.Count > 0)
                    scopes.Pop().Dispose();
            },
            capturedSetScalarIsExact ? capturedSetScalar : null);
    }

    public bool TryRecordNestedFieldResult(DxpFieldEvalResult result)
    {
        string? fieldType = new DxpFieldParser().Parse(InstructionText ?? string.Empty).Ast.FieldType;
        if (!string.Equals(fieldType, "SET", StringComparison.OrdinalIgnoreCase))
            return false;

        _nestedSetResultCount++;
        var currentParse = new DxpFieldParser().Parse(InstructionText ?? string.Empty);
        bool onlySetNamePrecedesResult = currentParse.Ast.ArgumentsText != null &&
            DxpFieldTokenization.TokenizeArgs(currentParse.Ast.ArgumentsText).Count == 1;
        _capturedSetScalarIsExact = _nestedSetResultCount == 1 &&
            onlySetNamePrecedesResult && result.Value.HasValue;
        _capturedSetScalar = _capturedSetScalarIsExact ? result.Value : null;

        string value = result.Text ?? string.Empty;
        bool needsSpace = !string.IsNullOrEmpty(InstructionText) &&
            !char.IsWhiteSpace(InstructionText![InstructionText.Length - 1]);
        string escaped = value.Replace("\"", "\\\"");
        AppendInstructionText((needsSpace ? " " : string.Empty) + $"\"{escaped}\"");
        return true;
    }

    private sealed class FieldEvent
    {
        private FieldEvent(FieldEventKind kind, object? data1 = null, object? data2 = null)
        {
            Kind = kind;
            Data1 = data1;
            Data2 = data2;
        }

        public FieldEventKind Kind { get; }
        public object? Data1 { get; }
        public object? Data2 { get; }

        public static FieldEvent Instruction(FieldCode instr, string text) => new(FieldEventKind.Instruction, instr, text);
        public static FieldEvent Separate(FieldChar separate) => new(FieldEventKind.Separate, separate);
        public static FieldEvent End(FieldChar end) => new(FieldEventKind.End, end);
        public static FieldEvent RunBegin(Run run) => new(FieldEventKind.RunBegin, run);
        public static FieldEvent RunEnd() => new(FieldEventKind.RunEnd);
        public static FieldEvent HyperlinkBegin(Hyperlink link, DxpLinkAnchor? target) => new(FieldEventKind.HyperlinkBegin, link, target);
        public static FieldEvent HyperlinkEnd() => new(FieldEventKind.HyperlinkEnd);
        public static FieldEvent Text(Text text) => new(FieldEventKind.Text, text);
        public static FieldEvent Break(Break br) => new(FieldEventKind.Break, br);
        public static FieldEvent Tab(TabChar tab) => new(FieldEventKind.Tab, tab);
        public static FieldEvent CarriageReturn(CarriageReturn cr) => new(FieldEventKind.CarriageReturn, cr);
        public static FieldEvent NoBreakHyphen(NoBreakHyphen nbh) => new(FieldEventKind.NoBreakHyphen, nbh);
        public static FieldEvent SimpleBegin(SimpleField fld) => new(FieldEventKind.SimpleBegin, fld);
        public static FieldEvent SimpleEnd() => new(FieldEventKind.SimpleEnd);
        public static FieldEvent ParagraphBegin(Paragraph paragraph, DxpIParagraphContext paragraphContext) => new(FieldEventKind.ParagraphBegin, paragraph, paragraphContext);
        public static FieldEvent ParagraphEnd() => new(FieldEventKind.ParagraphEnd);
        public static FieldEvent StructuredResult(DxpFieldNodeBuffer buffer) => new(FieldEventKind.StructuredResult, buffer);
        public static FieldEvent DeferredField(DxpDeferredField field) => new(FieldEventKind.DeferredField, field);

        public void Replay(DxpIVisitor visitor, DxpIDocumentContext d, Stack<IDisposable> scopes)
        {
            switch (Kind)
            {
                case FieldEventKind.Instruction:
                    visitor.VisitComplexFieldInstruction((FieldCode)Data1!, (string)Data2!, d);
                    break;
                case FieldEventKind.Separate:
                    visitor.VisitComplexFieldSeparate((FieldChar)Data1!, d);
                    break;
                case FieldEventKind.End:
                    visitor.VisitComplexFieldEnd((FieldChar)Data1!, d);
                    break;
                case FieldEventKind.RunBegin:
                    scopes.Push(visitor.VisitRunBegin((Run)Data1!, d));
                    break;
                case FieldEventKind.RunEnd:
                    if (scopes.Count > 0)
                        scopes.Pop().Dispose();
                    break;
                case FieldEventKind.HyperlinkBegin:
                    scopes.Push(visitor.VisitHyperlinkBegin((Hyperlink)Data1!, (DxpLinkAnchor?)Data2, d));
                    break;
                case FieldEventKind.HyperlinkEnd:
                    if (scopes.Count > 0)
                        scopes.Pop().Dispose();
                    break;
                case FieldEventKind.Text:
                    visitor.VisitText((Text)Data1!, d);
                    break;
                case FieldEventKind.Break:
                    visitor.VisitBreak((Break)Data1!, d);
                    break;
                case FieldEventKind.Tab:
                    visitor.VisitTab((TabChar)Data1!, d);
                    break;
                case FieldEventKind.CarriageReturn:
                    visitor.VisitCarriageReturn((CarriageReturn)Data1!, d);
                    break;
                case FieldEventKind.NoBreakHyphen:
                    visitor.VisitNoBreakHyphen((NoBreakHyphen)Data1!, d);
                    break;
                case FieldEventKind.SimpleBegin:
                    scopes.Push(visitor.VisitSimpleFieldBegin((SimpleField)Data1!, d));
                    break;
                case FieldEventKind.SimpleEnd:
                    if (scopes.Count > 0)
                        scopes.Pop().Dispose();
                    break;
                case FieldEventKind.ParagraphBegin:
                    scopes.Push(visitor.VisitParagraphBegin((Paragraph)Data1!, d, (DxpIParagraphContext)Data2!));
                    break;
                case FieldEventKind.ParagraphEnd:
                    if (scopes.Count > 0)
                        scopes.Pop().Dispose();
                    break;
                case FieldEventKind.StructuredResult:
                    if (visitor is IDxpStructuredFieldResultSink sink)
                        sink.TryRecordStructuredFieldResult((DxpFieldNodeBuffer)Data1!);
                    break;
                case FieldEventKind.DeferredField:
                    if (visitor is IDxpDeferredFieldSink deferredSink)
                        deferredSink.TryRecordDeferredField((DxpDeferredField)Data1!, d);
                    break;
            }
        }
    }

    private enum FieldEventKind
    {
        Instruction,
        Separate,
        End,
        RunBegin,
        RunEnd,
        HyperlinkBegin,
        HyperlinkEnd,
        Text,
        Break,
        Tab,
        CarriageReturn,
        NoBreakHyphen,
        SimpleBegin,
        SimpleEnd,
        ParagraphBegin,
        ParagraphEnd,
        StructuredResult,
        DeferredField
    }
}

internal interface IDxpNestedFieldResultSink
{
    bool TryRecordNestedFieldResult(DxpFieldEvalResult result);
}

internal sealed class DxpDeferredField
{
    private readonly Action<DxpIVisitor, DxpIDocumentContext> _replay;

    public DxpDeferredField(
        string instructionText,
        DxpFieldExpression expression,
        Action<DxpIVisitor, DxpIDocumentContext> replay,
        DxpFieldValue? capturedScalar = null)
    {
        InstructionText = instructionText;
        Expression = expression;
        _replay = replay;
        CapturedScalar = capturedScalar;
    }

    public string InstructionText { get; }
    public DxpFieldExpression Expression { get; }
    public DxpFieldValue? CapturedScalar { get; }
    public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => _replay(visitor, context);
}
