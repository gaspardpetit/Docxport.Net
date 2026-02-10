using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields;
using DocxportNet.Fields.Frames;
using DocxportNet.Middleware;
using Microsoft.Extensions.Logging;
using System;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpEvalGenericFieldFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    public bool SuppressContent { get; set; }
    public bool Evaluated { get; set; }
    public bool SeenSeparate { get; set; }
    public bool InResult { get; set; }
    public string? InstructionText { get; set; }
    public bool InstructionEmitted { get; set; }
    public RunProperties? CodeRunProperties { get; set; }
    public Run? CodeRun { get; set; }
    public List<Run?>? CachedResultRuns { get; set; }
    public List<RunProperties?>? CachedResultRunProperties { get; set; }
    public DxpIFCaptureState? IfState { get; set; }

	public override DxpIVisitor Next { get; }
	public DxpFieldEvalContext EvalContext { get; }

    private readonly DxpEvalFieldMode _mode;
    private readonly ILogger? _logger;

    public DxpEvalGenericFieldFrame(DxpIVisitor next, DxpFieldEval eval, DxpFieldEvalContext evalContext, ILogger? logger, DxpEvalFieldMode mode)
        : base()
    {
		Next = next ?? throw new ArgumentNullException(nameof(next));
		EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
        _mode = mode;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || InResult)
            return;
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!SeenSeparate)
        {
            SeenSeparate = true;
            InResult = true;
        }

        if (_mode == DxpEvalFieldMode.Evaluate)
            EmitUnsupported(d);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (_mode == DxpEvalFieldMode.Evaluate)
            EmitUnsupported(d);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (_mode != DxpEvalFieldMode.Cache)
            return;
        if (!InResult || SuppressContent || EvalContext.FieldDepth > 1)
            return;

        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("GenericCachedText: text='{Text}' depth={Depth}", text, EvalContext.FieldDepth);

        if (string.IsNullOrEmpty(text))
            return;

        Next.VisitComplexFieldCachedResultText(text, d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        if (_mode == DxpEvalFieldMode.Evaluate)
            EmitUnsupported(d);
        return DxpDisposable.Empty;
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
        => false;

    private void EmitUnsupported(DxpIDocumentContext d)
    {
        if (Evaluated)
            return;
        Evaluated = true;
        SuppressContent = true;

        var instruction = string.IsNullOrWhiteSpace(InstructionText) ? " " : InstructionText!;
        var text = DxpFieldEvalRules.GetEvaluationErrorText(new DxpFieldParser(), instruction);
        var t = new Text(text);

        var run = new Run();
		using (Next.VisitRunBegin(run, d))
			Next.VisitText(t, d);
    }
}
