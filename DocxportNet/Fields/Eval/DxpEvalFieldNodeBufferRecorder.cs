using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Visitors;
using DocxportNet.Walker;

namespace DocxportNet.Fields.Eval;

internal sealed class DxpEvalFieldNodeBufferRecorder : DxpVisitor
{
    private readonly Stack<DxpFieldNodeBuffer> _stack = new();
    private readonly Stack<Run?> _runStack = new();

    public DxpEvalFieldNodeBufferRecorder() : base(null)
    {
    }

    public void Reset(DxpFieldNodeBuffer root)
    {
        _stack.Clear();
        _runStack.Clear();
        _stack.Push(root);
        _runStack.Push(null);
    }

    private DxpFieldNodeBuffer Current => _stack.Peek();
    private Run? CurrentRun => _runStack.Peek();

    public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
    {
        var run = DxpRunCloner.CloneRunWithParagraphAncestor(r);
        var child = Current.BeginRun(run);
        _stack.Push(child);
        _runStack.Push(run);
        return DxpDisposable.Create(() => {
            _stack.Pop();
            _runStack.Pop();
        });
    }

    public override void VisitText(Text t, DxpIDocumentContext d)
    {
        Current.AddText(t.Text);
        var run = CurrentRun;
        if (run != null)
        {
            var text = new Text(t.Text);
            if (DxpFieldEvalMiddleware.NeedsPreserveSpace(t.Text))
                text.Space = SpaceProcessingModeValues.Preserve;
            run.AppendChild(text);
        }
    }
    public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d) => Current.AddDeletedText(dt.Text);
    public override void VisitBreak(Break b, DxpIDocumentContext d) => Current.AddBreak();
    public override void VisitTab(TabChar tab, DxpIDocumentContext d) => Current.AddTab();
    public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => Current.AddCarriageReturn();
    public override void VisitNoBreakHyphen(NoBreakHyphen nbh, DxpIDocumentContext d) => Current.AddNoBreakHyphen();

    public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
    {
        var clone = (Hyperlink)link.CloneNode(false);
        var child = Current.BeginHyperlink(clone, target);
        _stack.Push(child);
        return DxpDisposable.Create(() => _stack.Pop());
    }
}
