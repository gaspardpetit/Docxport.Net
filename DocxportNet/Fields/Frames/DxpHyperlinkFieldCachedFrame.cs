using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;

namespace DocxportNet.Fields.Frames;

internal sealed class DxpHyperlinkFieldCachedFrame : DxpFieldCachedFrameBase
{
    private readonly DxpIVisitor _linkVisitor;
    private readonly string _instruction;

    public DxpHyperlinkFieldCachedFrame(DxpIVisitor next, DxpIVisitor linkVisitor, string instruction)
        : base(next, instruction)
    {
        _linkVisitor = linkVisitor;
        _instruction = instruction;
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext context)
    {
        BeginLink(separate, context);
        base.VisitComplexFieldSeparate(separate, context);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext context)
    {
        base.VisitComplexFieldEnd(end, context);
        _linkVisitor.VisitComplexFieldEnd(end, context);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField field, DxpIDocumentContext context)
    {
        BeginLink(new FieldChar { FieldCharType = FieldCharValues.Separate }, context);
        var inner = base.VisitSimpleFieldBegin(field, context);
        return DxpDisposable.Create(() => {
            inner.Dispose();
            _linkVisitor.VisitComplexFieldEnd(new FieldChar { FieldCharType = FieldCharValues.End }, context);
        });
    }

    private void BeginLink(FieldChar separate, DxpIDocumentContext context)
    {
        _linkVisitor.VisitComplexFieldBegin(new FieldChar { FieldCharType = FieldCharValues.Begin }, context);
        _linkVisitor.VisitComplexFieldInstruction(new FieldCode(_instruction), _instruction, context);
        _linkVisitor.VisitComplexFieldSeparate(separate, context);
    }
}
