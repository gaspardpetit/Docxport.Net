using DocxportNet.API;
using DocxportNet.Middleware;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpPassthroughFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    private readonly DxpIVisitor _next;
    private bool _beginEmitted;

    internal FieldChar? BeginMarker { get; set; }

    public DxpPassthroughFieldEvalFrame(DxpIVisitor next) => _next = next;

    public override DxpIVisitor Next => _next;

    public override void VisitComplexFieldInstruction(FieldCode instruction, string text, DxpIDocumentContext context)
    {
        // A single field instruction can span several styled instrText runs.
        if (!_beginEmitted)
        {
            EmitFieldChar(BeginMarker ??
                new FieldChar { FieldCharType = FieldCharValues.Begin }, context);
            _beginEmitted = true;
        }
        var sourceRun = instruction.Parent as Run;
        var run = new Run();
        if (sourceRun?.RunProperties is { } properties)
            run.RunProperties = (RunProperties)properties.CloneNode(true);
        using (_next.VisitRunBegin(run, context))
            _next.VisitComplexFieldInstruction(instruction, text, context);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext context)
    {
        using (_next.VisitRunBegin(CloneRunShell(separate), context))
            _next.VisitComplexFieldSeparate(separate, context);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext context)
        => _next.VisitComplexFieldEnd(end, context);

    private void EmitFieldChar(FieldChar marker, DxpIDocumentContext context)
    {
        using (_next.VisitRunBegin(CloneRunShell(marker), context))
            _next.VisitComplexFieldBegin(marker, context);
    }

    private static Run CloneRunShell(FieldChar marker)
    {
        var run = new Run();
        if ((marker.Parent as Run)?.RunProperties is { } properties)
            run.RunProperties = (RunProperties)properties.CloneNode(true);
        return run;
    }
}
