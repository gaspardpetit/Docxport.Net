using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Eval;
using DocxportNet.Fields.Formatting;
using DocxportNet.Middleware;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using Microsoft.Extensions.Logging;
using System.Text;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpDocVariableFieldCachedFrame : DxpMiddleware, DxpIFieldEvalFrame
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


    private readonly ILogger? _logger;

    public DxpDocVariableFieldCachedFrame(DxpIVisitor next, DxpFieldEvalContext evalContext, ILogger? logger)
        : base()
    {
		Next = next ?? throw new ArgumentNullException(nameof(next));
        EvalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
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
            if (ReferenceEquals(EvalContext.ActiveIfFrame, this))
                EvalContext.ActiveIfFrame = null;
        }
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (Evaluated)
            return;
        if (_logger?.IsEnabled(LogLevel.Debug) == true)
            _logger.LogDebug("CachedText: text='{Text}' frame={Frame} suppress={Suppress} outerSuppress={OuterSuppress} depth={Depth}",
                text,
                GetType().Name,
                SuppressContent,
                EvalContext.OuterFrame?.SuppressContent == true,
                EvalContext.FieldDepth);

        if (EvalContext.OuterFrame?.SuppressContent == true)
            return;
        if (EvalContext.FieldDepth > 1)
            return;
        if (EvalContext.ActiveIfFrame != null)
            return;
        if (Evaluated)
            return;

        if (!InstructionEmitted && !string.IsNullOrWhiteSpace(InstructionText))
        {
            InstructionEmitted = true;
            Next.VisitComplexFieldInstruction(new FieldCode(), InstructionText!, d);
        }

        if (!ShouldForwardContent(d))
            return;

        Next.VisitComplexFieldCachedResultText(text, d);
        return;
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        if (!string.IsNullOrWhiteSpace(InstructionText))
            InstructionEmitted = true;

        return DxpDisposable.Empty;
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        return base.VisitRunBegin(r, d);
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
    {
        if (EvalContext.OuterFrame == null)
            return true;
        if (!EvalContext.OuterFrame.InResult)
            return false;
        if (EvalContext.OuterFrame.SuppressContent)
            return false;
        return EvalContext.FieldDepth <= 1;
    }
}
