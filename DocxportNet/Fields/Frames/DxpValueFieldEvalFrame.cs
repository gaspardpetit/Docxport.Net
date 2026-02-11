using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

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
    private List<Run?>? _cachedResultRuns;

    public override DxpIVisitor? Next { get; }

    protected DxpFieldEval Eval => _eval;
    protected DxpFieldEvalContext EvalContext => _eval.Context;
    protected string? InstructionText => _instructionText;

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
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        Evaluate(d);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        return;
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        _inCachedResult = true;
        return DxpDisposable.Create(() => {
            Evaluate(d);
            _inCachedResult = false;
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (_inCachedResult && DxpFieldEvalRules.HasRenderableContent(r))
        {
            _cachedResultRuns ??= new List<Run?>();
            _cachedResultRuns.Add(DxpRunCloner.CloneRunWithParagraphAncestor(r));
        }

        return DxpDisposable.Empty;
    }

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        return DxpDisposable.Empty;
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        return;
    }

    public override void VisitBreak(Break br, DxpIDocumentContext d)
    {
        return;
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        return;
    }

    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        return;
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
    {
        return;
    }

    protected virtual bool Evaluate(DxpIDocumentContext d)
    {
        if (string.IsNullOrWhiteSpace(_instructionText))
            return false;

        var result = _eval.EvalAsync(new DxpFieldInstruction(_instructionText!), d).GetAwaiter().GetResult();
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

        var parser = new DxpFieldParser();
        var parse = parser.Parse(_instructionText!);
        IReadOnlyList<IDxpFieldFormatSpec> formatSpecs = parse.Ast.FormatSpecs;
        return DxpFieldFrames.EmitTextWithMergeFormat(resultText, formatSpecs, _cachedResultRuns, _codeRun, d, Next, _logger);
    }
}
