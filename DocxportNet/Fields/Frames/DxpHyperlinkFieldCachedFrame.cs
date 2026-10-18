using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpHyperlinkFieldCachedFrame : DxpFieldCachedFrameBase
{
    private readonly string _instruction;

    public DxpHyperlinkFieldCachedFrame(DxpIVisitor next, string instruction)
        : base(next, instruction)
    {
        _instruction = instruction;
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext context)
    {
        Next?.VisitComplexFieldBegin(new FieldChar { FieldCharType = FieldCharValues.Begin }, context);
        Next?.VisitComplexFieldInstruction(new FieldCode(_instruction), _instruction, context);
        Next?.VisitComplexFieldSeparate(separate, context);
        base.VisitComplexFieldSeparate(separate, context);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext context)
    {
        base.VisitComplexFieldEnd(end, context);
        Next?.VisitComplexFieldEnd(end, context);
    }
}
