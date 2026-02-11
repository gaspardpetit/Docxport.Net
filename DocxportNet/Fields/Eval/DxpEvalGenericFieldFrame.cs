using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Frames;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpEvalGenericFieldFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    public bool Evaluated { get; set; }
    public string? InstructionText { get; set; }
    public Run? CodeRun { get; set; }

	public override DxpIVisitor Next { get; }
	public DxpFieldEvalContext EvalContext { get; }

    private readonly DxpFieldEval _eval;
    private readonly DxpEvalFieldMode _mode;
    private readonly ILogger? _logger;
    private bool _suppressContent;
    private bool _seenSeparate;
    private bool _inResult;
    private readonly List<FieldEvent> _events = new();
    private readonly Stack<IDisposable> _replayScopes = new();

    public DxpEvalGenericFieldFrame(
        DxpIVisitor next,
        DxpFieldEval eval,
        DxpFieldEvalContext evalContext,
        ILogger? logger,
        DxpEvalFieldMode mode,
        bool initialInResult = false,
        bool initialSeenSeparate = false,
        string? initialInstructionText = null)
        : base()
    {
		Next = next ?? throw new ArgumentNullException(nameof(next));
        _eval = eval ?? throw new ArgumentNullException(nameof(eval));
		EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
        _mode = mode;
        _inResult = initialInResult;
        _seenSeparate = initialSeenSeparate;
        InstructionText = initialInstructionText;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || _inResult)
            return;
        _events.Add(FieldEvent.Instruction(instr, text));
        AppendInstructionText(text);
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

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (_mode != DxpEvalFieldMode.Cache)
            return;
        if (!_inResult || _suppressContent || EvalContext.FieldDepth > 1)
            return;

        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("GenericCachedText: text='{Text}' depth={Depth}", text, EvalContext.FieldDepth);

        if (string.IsNullOrEmpty(text))
            return;
        _events.Add(FieldEvent.CachedResult(text));
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
            return;
        }
        _events.Add(FieldEvent.NoBreakHyphen(nbh));
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        return Next.VisitParagraphBegin(p, d, paragraph);
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
        => false;

    private void EmitUnsupported(DxpIDocumentContext d)
    {
        if (Evaluated)
            return;
        Evaluated = true;
        _suppressContent = true;

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
        DxpIFieldEvalFrame? delegateFrame = CreateDelegateFrame(InstructionText);
        if (delegateFrame == null)
        {
            if (_mode == DxpEvalFieldMode.Evaluate)
                EmitUnsupported(d);
            _events.Clear();
            return;
        }

        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug(
                "Generic.Replay: frame={Frame} mode={Mode} events={EventCount}",
                delegateFrame.GetType().Name,
                _mode,
                _events.Count);

        foreach (var ev in _events)
            ev.Replay(delegateFrame, d, _replayScopes);
        while (_replayScopes.Count > 0)
            _replayScopes.Pop().Dispose();
        _events.Clear();
    }

    internal void TryCaptureCodeRun(Run r)
    {
        if (CodeRun == null && !_inResult)
            CodeRun = DxpRunCloner.CloneRunWithParagraphAncestor(r);
    }

    private DxpIFieldEvalFrame? CreateDelegateFrame(string? instructionText)
    {
        if (string.IsNullOrWhiteSpace(instructionText))
            return null;
        if (DxpFieldEvalFrameFactory.IsSetInstruction(instructionText))
        {
			DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpSetFieldCachedFrame(EvalContext, _logger)
                : new DxpSetFieldEvalFrame(_eval, EvalContext, _logger, instructionText);
            return frame;
        }
        if (DxpFieldEvalFrameFactory.IsRefInstruction(instructionText))
        {
			DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpRefFieldCachedFrame(Next)
                : new DxpRefFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
            return frame;
        }
        if (DxpFieldEvalFrameFactory.IsDocVariableInstruction(instructionText))
        {
			DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpDocVariableFieldCachedFrame(Next)
                : new DxpDocVariableFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
            return frame;
        }
        if (DxpFieldEvalFrameFactory.IsIfInstruction(instructionText))
        {
			DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpIFFieldCachedFrame(Next)
                : new DxpIFFieldEvalFrame(Next, _eval, _logger, CodeRun);
            return frame;
        }
        if (DxpFieldEvalFrameFactory.IsAskInstruction(instructionText))
        {
            DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpAskFieldCachedFrame(Next)
                : new DxpAskFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
            return frame;
        }
        if (DxpFieldEvalFrameFactory.IsSkipIfInstruction(instructionText))
        {
            DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpSkipIfFieldCachedFrame(Next)
                : new DxpSkipIfFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
            return frame;
        }
        if (DxpFieldEvalFrameFactory.IsDocPropertyInstruction(instructionText) ||
            DxpFieldEvalFrameFactory.IsMergeFieldInstruction(instructionText) ||
            DxpFieldEvalFrameFactory.IsSeqInstruction(instructionText) ||
            DxpFieldEvalFrameFactory.IsDateTimeInstruction(instructionText) ||
            DxpFieldEvalFrameFactory.IsCompareInstruction(instructionText) ||
            DxpFieldEvalFrameFactory.IsDocumentMetricInstruction(instructionText))
        {
            if (DxpFieldEvalFrameFactory.IsDocPropertyInstruction(instructionText))
            {
                DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                    ? new DxpDocPropertyFieldCachedFrame(Next)
                    : new DxpDocPropertyFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
                return frame;
            }
            if (DxpFieldEvalFrameFactory.IsMergeFieldInstruction(instructionText))
            {
                DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                    ? new DxpMergeFieldCachedFrame(Next)
                    : new DxpMergeFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
                return frame;
            }
            if (DxpFieldEvalFrameFactory.IsSeqInstruction(instructionText))
            {
                DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                    ? new DxpSeqFieldCachedFrame(Next)
                    : new DxpSeqFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
                return frame;
            }
            if (DxpFieldEvalFrameFactory.IsDateTimeInstruction(instructionText))
            {
                DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                    ? new DxpDateTimeFieldCachedFrame(Next)
                    : new DxpDateTimeFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
                return frame;
            }
            if (DxpFieldEvalFrameFactory.IsCompareInstruction(instructionText))
            {
                DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                    ? new DxpCompareFieldCachedFrame(Next)
                    : new DxpCompareFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
                return frame;
            }
            if (DxpFieldEvalFrameFactory.IsDocumentMetricInstruction(instructionText))
            {
                DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                    ? new DxpSimpleFieldCachedFrame(Next)
                    : new DxpValueFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
                return frame;
            }
        }
        if (DxpFieldEvalFrameFactory.IsFormulaInstruction(instructionText))
        {
            DxpIFieldEvalFrame frame = _mode == DxpEvalFieldMode.Cache
                ? new DxpFormulaFieldCachedFrame(Next)
                : new DxpFormulaFieldEvalFrame(Next, _eval, _logger, instructionText, CodeRun);
            return frame;
        }
        return null;
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
        public static FieldEvent CachedResult(string text) => new(FieldEventKind.CachedResult, text);
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
                case FieldEventKind.CachedResult:
                    visitor.VisitComplexFieldCachedResultText((string)Data1!, d);
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
            }
        }
    }

    private enum FieldEventKind
    {
        Instruction,
        Separate,
        End,
        CachedResult,
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
        SimpleEnd
    }
}
