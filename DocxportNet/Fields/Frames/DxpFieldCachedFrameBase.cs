using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Middleware;

namespace DocxportNet.Fields.Frames;

internal abstract class DxpFieldCachedFrameBase : DxpMiddleware, DxpIFieldEvalFrame
{
    private bool _inCachedResult;
    private readonly string? _instructionText;
    private bool _pushedFieldContext;

    public override DxpIVisitor? Next { get; }

    protected DxpFieldCachedFrameBase(DxpIVisitor? next, string? instructionText = null)
        : base()
    {
        Next = next;
        _instructionText = instructionText;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || _inCachedResult)
            return;
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        _inCachedResult = true;
        PushFieldContext(d);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        _inCachedResult = false;
        PopFieldContext(d);
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (_inCachedResult && Next != null && !string.IsNullOrEmpty(text))
            Next.VisitComplexFieldCachedResultText(text, d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        _inCachedResult = true;
        PushFieldContext(d);
        return DxpDisposable.Create(() => {
            _inCachedResult = false;
            PopFieldContext(d);
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
        => _inCachedResult && Next != null ? Next.VisitRunBegin(r, d) : DxpDisposable.Empty;

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
        => _inCachedResult && Next != null ? Next.VisitHyperlinkBegin(link, target, d) : DxpDisposable.Empty;

    public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
        => _inCachedResult && Next != null ? Next.VisitBlockBegin(child, d) : DxpDisposable.Empty;

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
        => _inCachedResult && Next != null ? Next.VisitParagraphBegin(p, d, paragraph) : DxpDisposable.Empty;

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (_inCachedResult && Next != null)
            Next.VisitComplexFieldCachedResultText(t.Text, d);
    }

    public override void VisitBreak(Break br, DxpIDocumentContext d)
    {
        if (_inCachedResult && Next != null)
            Next.VisitBreak(br, d);
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        if (_inCachedResult && Next != null)
            Next.VisitTab(tab, d);
    }

    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        if (_inCachedResult && Next != null)
            Next.VisitCarriageReturn(cr, d);
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
    {
        if (_inCachedResult && Next != null)
            Next.VisitNoBreakHyphen(nbh, d);
    }

    private void PushFieldContext(DxpIDocumentContext d)
    {
        if (_pushedFieldContext)
            return;

        d.CurrentFields.FieldStack.Push(new FieldFrame {
            SeenSeparate = true,
            InResult = true,
            InstructionText = _instructionText
        });
        _pushedFieldContext = true;
    }

    private void PopFieldContext(DxpIDocumentContext d)
    {
        if (!_pushedFieldContext)
            return;

        if (d.CurrentFields.FieldStack.Count > 0)
            d.CurrentFields.FieldStack.Pop();
        _pushedFieldContext = false;
    }
}
