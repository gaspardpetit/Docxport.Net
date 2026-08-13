using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal class DxpValueFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    private readonly ILogger? _logger;
    private readonly DxpFieldEval _eval;
    private readonly bool _emitResult;
    private readonly bool _emitErrorOnFailure;

    private bool _inCachedResult;
    private string? _instructionText;
    private Run? _codeRun;
    private readonly DxpEvalFieldNodeBufferRecorder _recorder = new();
    private DxpFieldNodeBuffer? _cachedResultBuffer;

    public override DxpIVisitor? Next { get; }

    protected DxpFieldEval Eval => _eval;
    protected DxpFieldEvalContext EvalContext => _eval.Context;
    protected string? InstructionText => _instructionText;
    protected ILogger? Logger => _logger;
    protected DxpFieldNodeBuffer? CachedResultBuffer => _cachedResultBuffer;

    public DxpValueFieldEvalFrame(
        DxpIVisitor? next,
        DxpFieldEval eval,
        ILogger? logger,
        string? instructionText,
        Run? codeRun = null,
        bool emitResult = true,
        bool emitErrorOnFailure = false)
        : base()
    {
        Next = next;
        _eval = eval ?? throw new ArgumentNullException(nameof(eval));
        _logger = logger;
        _instructionText = instructionText;
        _codeRun = codeRun;
        _emitResult = emitResult;
        _emitErrorOnFailure = emitErrorOnFailure;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || _inCachedResult)
            return;
        if (_codeRun == null && instr.Parent is Run instrRun)
            _codeRun = DxpRunCloner.CloneRunWithParagraphAncestor(instrRun);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        _inCachedResult = true;
        BeginCachedCapture();
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        Evaluate(d);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (_inCachedResult && !string.IsNullOrEmpty(text))
            _recorder.VisitText(new Text(text), d);
        return;
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        _inCachedResult = true;
        BeginCachedCapture();
        return DxpDisposable.Create(() => {
            Evaluate(d);
            _inCachedResult = false;
        });
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
        => _inCachedResult ? _recorder.VisitParagraphBegin(p, d, paragraph) : DxpDisposable.Empty;

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (!_inCachedResult)
            return DxpDisposable.Empty;

        return _recorder.VisitRunBegin(r, d);
    }

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        if (!_inCachedResult)
            return DxpDisposable.Empty;

        return _recorder.VisitHyperlinkBegin(link, target, d);
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (_inCachedResult)
            _recorder.VisitText(t, d);
        return;
    }

    public override void VisitBreak(Break br, DxpIDocumentContext d)
    {
        if (_inCachedResult)
            _recorder.VisitBreak(br, d);
        return;
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        if (_inCachedResult)
            _recorder.VisitTab(tab, d);
        return;
    }

    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        if (_inCachedResult)
            _recorder.VisitCarriageReturn(cr, d);
        return;
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
    {
        if (_inCachedResult)
            _recorder.VisitNoBreakHyphen(nbh, d);
        return;
    }

    protected virtual bool Evaluate(DxpIDocumentContext d)
    {
        if (string.IsNullOrWhiteSpace(_instructionText))
            return false;

        var cachedResultText = _cachedResultBuffer?.ToPlainText();
        var result = _eval.EvalAsync(new DxpFieldInstruction(_instructionText!, cachedResultText), d).GetAwaiter().GetResult();
        if (!_emitResult)
            return true;
        if (result.Status == DxpFieldEvalStatus.Skipped)
            return true;

        string? resultText = result.Text;
        if (resultText == null)
        {
            if (!_emitErrorOnFailure)
                return true;
            resultText = DxpFieldEvalRules.GetEvaluationErrorText(_instructionText!);
        }

        if (EvalContext.FieldDepth > 1 && Next is IDxpNestedFieldResultSink nestedSink &&
            nestedSink.TryRecordNestedFieldResult(result))
            return true;

        var parser = new DxpFieldParser();
        var parse = parser.Parse(_instructionText!);
        IReadOnlyList<IDxpFieldFormatSpec> formatSpecs = parse.Ast.FormatSpecs;
        return DxpFieldFrames.EmitTextWithMergeFormat(resultText, formatSpecs, _cachedResultBuffer, _codeRun, d, Next, _logger);
    }

    private void BeginCachedCapture()
    {
        _cachedResultBuffer = new DxpFieldNodeBuffer();
        _recorder.Reset(_cachedResultBuffer);
    }
}
