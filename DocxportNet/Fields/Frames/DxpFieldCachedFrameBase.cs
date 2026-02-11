using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Middleware;

namespace DocxportNet.Fields.Frames;

internal abstract class DxpFieldCachedFrameBase : DxpMiddleware, DxpIFieldEvalFrame
{
    private bool _inCachedResult;

    public override DxpIVisitor? Next { get; }

    protected DxpFieldCachedFrameBase(DxpIVisitor? next)
        : base()
    {
        Next = next;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || _inCachedResult)
            return;
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        _inCachedResult = true;
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        _inCachedResult = false;
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (_inCachedResult)
            Next?.VisitComplexFieldCachedResultText(text, d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        _inCachedResult = true;
        return DxpDisposable.Create(() => {
            _inCachedResult = false;
        });
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
        => _inCachedResult ? Next?.VisitRunBegin(r, d)?? DxpDisposable.Empty : DxpDisposable.Empty;

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
        => _inCachedResult ? Next?.VisitHyperlinkBegin(link, target, d)?? DxpDisposable.Empty : DxpDisposable.Empty;

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (_inCachedResult)
			Next?.VisitText(t, d);
    }

    public override void VisitBreak(Break br, DxpIDocumentContext d)
    {
        if (_inCachedResult)
			Next?.VisitBreak(br, d);
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        if (_inCachedResult)
			Next?.VisitTab(tab, d);
    }

    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        if (_inCachedResult)
			Next?.VisitCarriageReturn(cr, d);
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
    {
        if (_inCachedResult)
			Next?.VisitNoBreakHyphen(nbh, d);
    }
}
