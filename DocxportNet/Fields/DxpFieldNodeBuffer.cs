using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using DocxportNet.Walker.Context;
using System.Linq;
using System.Text;

namespace DocxportNet.Fields;

public sealed class DxpFieldNodeBuffer
{
    private interface IReplayNode
    {
        void Replay(DxpIVisitor visitor, DxpIDocumentContext context);
        OpenXmlElement CreateElement(bool forRoot);
        void AppendText(StringBuilder sb);
    }

    private sealed class TextNode : IReplayNode
    {
        private readonly string _text;

        public TextNode(string text)
        {
            _text = text;
        }

        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
        {
            var t = new Text(_text);
            if (NeedsPreserveSpace(_text))
                t.Space = SpaceProcessingModeValues.Preserve;
            visitor.VisitText(t, context);
        }

        public OpenXmlElement CreateElement(bool forRoot)
        {
            var t = new Text(_text);
            if (NeedsPreserveSpace(_text))
                t.Space = SpaceProcessingModeValues.Preserve;
            return t;
        }

        public void AppendText(StringBuilder sb) => sb.Append(_text);
    }

    private sealed class DeletedTextNode : IReplayNode
    {
        private readonly string _text;

        public DeletedTextNode(string text)
        {
            _text = text;
        }

        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
        {
            visitor.VisitDeletedText(new DeletedText(_text), context);
        }

        public OpenXmlElement CreateElement(bool forRoot) => new DeletedText(_text);

        public void AppendText(StringBuilder sb) => sb.Append(_text);
    }

    private sealed class BreakNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitBreak(new Break(), context);
        public OpenXmlElement CreateElement(bool forRoot) => new Break();
        public void AppendText(StringBuilder sb) => sb.Append('\n');
    }

    private sealed class TabNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitTab(new TabChar(), context);
        public OpenXmlElement CreateElement(bool forRoot) => new TabChar();
        public void AppendText(StringBuilder sb) => sb.Append('\t');
    }

    private sealed class CarriageReturnNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitCarriageReturn(new CarriageReturn(), context);
        public OpenXmlElement CreateElement(bool forRoot) => new CarriageReturn();
        public void AppendText(StringBuilder sb) => sb.Append('\n');
    }

    private sealed class NoBreakHyphenNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitNoBreakHyphen(new NoBreakHyphen(), context);
        public OpenXmlElement CreateElement(bool forRoot) => new NoBreakHyphen();
        public void AppendText(StringBuilder sb) => sb.Append('-');
    }

    private sealed class ParagraphNode : IReplayNode
    {
        private readonly Paragraph _paragraph;
        private readonly DxpFieldNodeBuffer _children;

        public ParagraphNode(Paragraph paragraph, DxpFieldNodeBuffer children)
        {
            _paragraph = paragraph;
            _children = children;
        }

        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
        {
            using (visitor.VisitBlockBegin(_paragraph, context))
            using (visitor.VisitParagraphBegin(_paragraph, context, context.CurrentParagraph))
                _children.Replay(visitor, context);
        }

        public OpenXmlElement CreateElement(bool forRoot)
        {
            var paragraph = (Paragraph)_paragraph.CloneNode(false);
            if (_paragraph.ParagraphProperties != null && paragraph.ParagraphProperties == null)
                paragraph.ParagraphProperties = (ParagraphProperties)_paragraph.ParagraphProperties.CloneNode(true);

            foreach (var child in _children.CreateElements(forRoot: false))
                paragraph.AppendChild(child);

            return paragraph;
        }

        public void AppendText(StringBuilder sb) => _children.AppendText(sb);
    }

    private sealed class RunNode : IReplayNode
    {
        private readonly Run _run;
        private readonly DxpFieldNodeBuffer _children;

        public RunNode(Run run, DxpFieldNodeBuffer children)
        {
            _run = run;
            _children = children;
        }

        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
        {
            using (visitor.VisitRunBegin(_run, context))
                _children.Replay(visitor, context);
        }

        public OpenXmlElement CreateElement(bool forRoot)
        {
            var run = (Run)_run.CloneNode(false);
            if (_run.RunProperties != null && run.RunProperties == null)
                run.RunProperties = (RunProperties)_run.RunProperties.CloneNode(true);

            foreach (var child in _children.CreateElements(forRoot: false))
                run.AppendChild(child);

            if (forRoot && _run.Parent is Paragraph paragraph)
            {
                var paraClone = (Paragraph)paragraph.CloneNode(false);
                if (paragraph.ParagraphProperties != null && paraClone.ParagraphProperties == null)
                    paraClone.ParagraphProperties = (ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true);
                paraClone.AppendChild(run);
                return paraClone;
            }
            else if (forRoot)
            {
                var paragraphWrapper = new Paragraph();
                paragraphWrapper.AppendChild(run);
                return paragraphWrapper;
            }

            return run;
        }

        public void AppendText(StringBuilder sb) => _children.AppendText(sb);

        public RunProperties? CloneRunProperties()
        {
            if (_run.RunProperties == null)
                return null;
            return (RunProperties)_run.RunProperties.CloneNode(true);
        }

        public string GetText() => _children.ToPlainText();

        public bool TryGetFirstRunProperties(out RunProperties? props)
        {
            props = CloneRunProperties();
            return true;
        }
    }

    private sealed class HyperlinkNode : IReplayNode
    {
        private readonly Hyperlink _link;
        private readonly DxpLinkAnchor? _target;
        private readonly DxpFieldNodeBuffer _children;

        public HyperlinkNode(Hyperlink link, DxpLinkAnchor? target, DxpFieldNodeBuffer children)
        {
            _link = link;
            _target = target;
            _children = children;
        }

        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
        {
            using (visitor.VisitHyperlinkBegin(_link, _target, context))
                _children.Replay(visitor, context);
        }

        public OpenXmlElement CreateElement(bool forRoot)
        {
            var link = (Hyperlink)_link.CloneNode(false);
            foreach (var child in _children.CreateElements(forRoot: false))
                link.AppendChild(child);
            if (forRoot)
            {
                if (_link.Parent is Paragraph paragraph)
                {
                    var paraClone = (Paragraph)paragraph.CloneNode(false);
                    if (paragraph.ParagraphProperties != null && paraClone.ParagraphProperties == null)
                        paraClone.ParagraphProperties = (ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true);
                    paraClone.AppendChild(link);
                    return paraClone;
                }
                else
                {
                    var paragraphWrapper = new Paragraph();
                    paragraphWrapper.AppendChild(link);
                    return paragraphWrapper;
                }
            }
            return link;
        }

        public void AppendText(StringBuilder sb) => _children.AppendText(sb);

        public bool TryGetFirstRunProperties(out RunProperties? props) => _children.TryGetFirstRunProperties(out props);
    }

    private readonly List<IReplayNode> _nodes;

    public DxpFieldNodeBuffer() : this(new List<IReplayNode>())
    {
    }

    private DxpFieldNodeBuffer(List<IReplayNode> nodes)
    {
        _nodes = nodes;
    }

    public static DxpFieldNodeBuffer FromText(string text)
    {
        var buffer = new DxpFieldNodeBuffer();
        buffer.AddRunText(text);
        return buffer;
    }

    public void Replay(DxpIVisitor visitor, DxpIDocumentContext context)
    {
        if (context is DxpDocumentContext docContext)
        {
            bool hasParagraphRoots = _nodes.Any(static n => n is ParagraphNode);

            if (!hasParagraphRoots)
            {
                var paragraph = CreateSyntheticParagraph(_nodes, docContext);
                foreach (var element in CreateElements(forRoot: false))
                    paragraph.AppendChild(element);
                docContext.Walker.WalkSyntheticFieldInlineContent(paragraph, docContext, visitor);
                return;
            }

            var pendingInline = new List<IReplayNode>();

            void FlushInline()
            {
                if (pendingInline.Count == 0)
                    return;

                Paragraph paragraph;
                var firstRoot = pendingInline[0].CreateElement(forRoot: true);
                if (firstRoot is Paragraph sourceParagraph)
                {
                    paragraph = (Paragraph)sourceParagraph.CloneNode(false);
                    if (sourceParagraph.ParagraphProperties != null && paragraph.ParagraphProperties == null)
                        paragraph.ParagraphProperties = (ParagraphProperties)sourceParagraph.ParagraphProperties.CloneNode(true);
                }
                else if (firstRoot.Ancestors<Paragraph>().FirstOrDefault() is Paragraph ancestorParagraph)
                {
                    paragraph = (Paragraph)ancestorParagraph.CloneNode(false);
                    if (ancestorParagraph.ParagraphProperties != null && paragraph.ParagraphProperties == null)
                        paragraph.ParagraphProperties = (ParagraphProperties)ancestorParagraph.ParagraphProperties.CloneNode(true);
                }
                else
                {
                    paragraph = new Paragraph();
                }

                foreach (var inline in pendingInline)
                    paragraph.AppendChild(inline.CreateElement(forRoot: false));

                docContext.Walker.WalkSyntheticFieldElement(paragraph, docContext, visitor);
                pendingInline.Clear();
            }

            foreach (var node in _nodes)
            {
                if (node is ParagraphNode)
                {
                    FlushInline();
                    var element = node.CreateElement(forRoot: true);
                    docContext.Walker.WalkSyntheticFieldElement(element, docContext, visitor);
                    continue;
                }

                pendingInline.Add(node);
            }

            FlushInline();
            return;
        }

        foreach (var node in _nodes)
            node.Replay(visitor, context);
    }

    public bool IsEmpty => _nodes.Count == 0;

    public string ToPlainText()
    {
        var sb = new StringBuilder();
        AppendText(sb);
        return sb.ToString();
    }

    public bool TryGetFirstRunProperties(out RunProperties? props)
    {
        props = null;
        foreach (var node in _nodes)
        {
            if (node is RunNode runNode)
                return runNode.TryGetFirstRunProperties(out props);
            if (node is HyperlinkNode linkNode)
                return linkNode.TryGetFirstRunProperties(out props);
        }
        return false;
    }

    internal bool TryGetRunSegments(out List<(string text, RunProperties? props)> segments)
    {
        segments = new List<(string text, RunProperties? props)>();
        foreach (var node in _nodes)
        {
            if (node is HyperlinkNode or ParagraphNode)
                return false;
            if (node is RunNode runNode)
            {
                var text = runNode.GetText();
                if (string.IsNullOrEmpty(text))
                    continue;
                segments.Add((text, runNode.CloneRunProperties()));
                continue;
            }
        }
        return segments.Count > 0;
    }


    internal void AddText(string text) => _nodes.Add(new TextNode(text));
    internal void AddDeletedText(string text) => _nodes.Add(new DeletedTextNode(text));
    internal void AddBreak() => _nodes.Add(new BreakNode());
    internal void AddTab() => _nodes.Add(new TabNode());
    internal void AddCarriageReturn() => _nodes.Add(new CarriageReturnNode());
    internal void AddNoBreakHyphen() => _nodes.Add(new NoBreakHyphenNode());

    internal DxpFieldNodeBuffer BeginParagraph(Paragraph paragraph)
    {
        var child = new DxpFieldNodeBuffer();
        _nodes.Add(new ParagraphNode((Paragraph)paragraph.CloneNode(false), child));
        return child;
    }

    internal DxpFieldNodeBuffer BeginRun(Run run)
    {
        var child = new DxpFieldNodeBuffer();
        _nodes.Add(new RunNode(run, child));
        return child;
    }

    internal DxpFieldNodeBuffer BeginHyperlink(Hyperlink link, DxpLinkAnchor? target)
    {
        var child = new DxpFieldNodeBuffer();
        _nodes.Add(new HyperlinkNode(link, target, child));
        return child;
    }

    private void AddRunText(string text)
    {
        var run = new Run();
        var t = new Text(text);
        if (NeedsPreserveSpace(text))
            t.Space = SpaceProcessingModeValues.Preserve;
        run.AppendChild(t);
        var child = BeginRun(run);
        child.AddText(text);
    }

    private void AppendText(StringBuilder sb)
    {
        bool first = true;
        foreach (var node in _nodes)
        {
            if (!first && node is ParagraphNode)
                sb.AppendLine();
            node.AppendText(sb);
            first = false;
        }
    }

    private List<OpenXmlElement> CreateElements(bool forRoot)
        => _nodes.Select(n => n.CreateElement(forRoot)).ToList();

    private static Paragraph CreateSyntheticParagraph(IReadOnlyList<IReplayNode> nodes, DxpDocumentContext? docContext)
    {
        if (docContext?.CurrentParagraph?.Properties != null)
        {
            var fromContext = new Paragraph();
            fromContext.ParagraphProperties = (ParagraphProperties)docContext.CurrentParagraph.Properties.CloneNode(true);
            return fromContext;
        }

        if (nodes.Count > 0)
        {
            var firstRoot = nodes[0].CreateElement(forRoot: true);
            if (firstRoot is Paragraph sourceParagraph)
            {
                var cloned = (Paragraph)sourceParagraph.CloneNode(false);
                if (sourceParagraph.ParagraphProperties != null && cloned.ParagraphProperties == null)
                    cloned.ParagraphProperties = (ParagraphProperties)sourceParagraph.ParagraphProperties.CloneNode(true);
                return cloned;
            }

            if (firstRoot.Ancestors<Paragraph>().FirstOrDefault() is Paragraph ancestorParagraph)
            {
                var cloned = (Paragraph)ancestorParagraph.CloneNode(false);
                if (ancestorParagraph.ParagraphProperties != null && cloned.ParagraphProperties == null)
                    cloned.ParagraphProperties = (ParagraphProperties)ancestorParagraph.ParagraphProperties.CloneNode(true);
                return cloned;
            }
        }

        return new Paragraph();
    }

    private static bool NeedsPreserveSpace(string text)
    {
        if (text.Length == 0)
            return false;
        if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
            return true;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '\t' || ch == '\r' || ch == '\n')
                return true;
            if (ch == ' ' && i + 1 < text.Length && text[i + 1] == ' ')
                return true;
        }
        return false;
    }
}
