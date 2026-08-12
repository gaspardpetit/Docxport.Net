using DocxportNet.API;
using DocxportNet.Middleware;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpPassthroughFieldEvalFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    private readonly DxpIVisitor _next;

    public DxpPassthroughFieldEvalFrame(DxpIVisitor next) => _next = next;

    public override DxpIVisitor Next => _next;

    public override void VisitComplexFieldInstruction(FieldCode instruction, string text, DxpIDocumentContext context)
    {
        EmitFieldChar(FieldCharValues.Begin, context);
        using (_next.VisitRunBegin(new Run(), context))
            _next.VisitComplexFieldInstruction(instruction, text, context);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext context)
    {
        using (_next.VisitRunBegin(new Run(), context))
            _next.VisitComplexFieldSeparate(separate, context);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext context)
        => _next.VisitComplexFieldEnd(end, context);

    private void EmitFieldChar(FieldCharValues type, DxpIDocumentContext context)
    {
        using (_next.VisitRunBegin(new Run(), context))
            _next.VisitComplexFieldBegin(new FieldChar { FieldCharType = type }, context);
    }
}
