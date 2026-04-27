using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Fields.Frames;
using DocxportNet.Middleware;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpCachedFieldRouterFrame : DxpMiddleware, DxpIFieldEvalFrame
{
    private static readonly DxpCachedFieldFrameFactory FrameFactory = new();

    public override DxpIVisitor Next { get; }

    private readonly DxpFieldEvalContext _evalContext;
    private readonly ILogger? _logger;
    private string? _instructionText;
    private bool _seenSeparate;
    private bool _inResult;
    private DxpIFieldEvalFrame? _delegate;

    public DxpCachedFieldRouterFrame(
        DxpIVisitor next,
        DxpFieldEvalContext evalContext,
        ILogger? logger,
        bool initialInResult = false,
        bool initialSeenSeparate = false,
        string? initialInstructionText = null)
        : base()
    {
        Next = next ?? throw new ArgumentNullException(nameof(next));
        _evalContext = evalContext ?? throw new ArgumentNullException(nameof(evalContext));
        _logger = logger;
        _inResult = initialInResult;
        _seenSeparate = initialSeenSeparate;
        _instructionText = initialInstructionText;
    }

    public override void VisitComplexFieldInstruction(FieldCode instr, string text, DxpIDocumentContext d)
    {
        if (string.IsNullOrEmpty(text) || _inResult)
            return;
        AppendInstructionText(text);
    }

    public override void VisitComplexFieldSeparate(FieldChar separate, DxpIDocumentContext d)
    {
        if (!_seenSeparate)
        {
            _seenSeparate = true;
            _inResult = true;
        }

        _delegate = FrameFactory.Create(_instructionText, Next, _evalContext, _logger);
        _delegate.VisitComplexFieldSeparate(separate, d);
    }

    public override void VisitComplexFieldEnd(FieldChar end, DxpIDocumentContext d)
    {
        if (_delegate == null)
            return;

        _delegate.VisitComplexFieldEnd(end, d);
        _delegate = null;
    }

    public override void VisitComplexFieldCachedResultText(string text, DxpIDocumentContext d)
    {
        if (!_inResult || _delegate == null)
            return;

        _delegate.VisitComplexFieldCachedResultText(text, d);
    }

    public override IDisposable VisitSimpleFieldBegin(SimpleField fld, DxpIDocumentContext d)
    {
        _inResult = true;
        _seenSeparate = true;
        _delegate = FrameFactory.Create(_instructionText, Next, _evalContext, _logger);
        return _delegate.VisitSimpleFieldBegin(fld, d);
    }

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
            return _delegate.VisitRunBegin(r, d);
        return DxpDisposable.Empty;
    }

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
            return _delegate.VisitHyperlinkBegin(link, target, d);
        return DxpDisposable.Empty;
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
        {
            _delegate.VisitText(t, d);
            return;
        }

        if (!_inResult)
            AppendInstructionText(t.Text);
    }

    public override void VisitBreak(Break br, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
        {
            _delegate.VisitBreak(br, d);
            return;
        }

        if (!_inResult)
            AppendInstructionText("\n");
    }

    public override void VisitTab(TabChar tab, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
        {
            _delegate.VisitTab(tab, d);
            return;
        }

        if (!_inResult)
            AppendInstructionText("\t");
    }

    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
        {
            _delegate.VisitCarriageReturn(cr, d);
            return;
        }

        if (!_inResult)
            AppendInstructionText("\n");
    }

    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
        {
            _delegate.VisitNoBreakHyphen(nbh, d);
            return;
        }

        if (!_inResult)
            AppendInstructionText("-");
    }

    public override IDisposable VisitBlockBegin(OpenXmlElement child, DxpIDocumentContext d)
    {
        if (_delegate != null && _inResult)
            return _delegate.VisitBlockBegin(child, d);
        if (_inResult)
            return DxpDisposable.Empty;
        return Next.VisitBlockBegin(child, d);
    }

    public override IDisposable VisitParagraphBegin(Paragraph p, DxpIDocumentContext d, DxpIParagraphContext paragraph)
    {
        if (_delegate != null && _inResult)
            return _delegate.VisitParagraphBegin(p, d, paragraph);
        if (_inResult)
            return DxpDisposable.Empty;
        return Next.VisitParagraphBegin(p, d, paragraph);
    }

    protected override bool ShouldForwardContent(DxpIDocumentContext d)
        => false;

    private void AppendInstructionText(string text)
    {
        if (string.IsNullOrEmpty(text))
            return;
        _instructionText = _instructionText == null ? text : _instructionText + text;
    }
}
