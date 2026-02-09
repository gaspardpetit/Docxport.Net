using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using System.Text;

namespace DocxportNet.Fields;

public sealed class DxpFieldNodeBuffer
{
    private interface IReplayNode
    {
        void Replay(DxpIVisitor visitor, DxpIDocumentContext context);
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

        public void AppendText(StringBuilder sb) => sb.Append(_text);
    }

    private sealed class BreakNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitBreak(new Break(), context);
        public void AppendText(StringBuilder sb) => sb.Append('\n');
    }

    private sealed class TabNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitTab(new TabChar(), context);
        public void AppendText(StringBuilder sb) => sb.Append('\t');
    }

    private sealed class CarriageReturnNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitCarriageReturn(new CarriageReturn(), context);
        public void AppendText(StringBuilder sb) => sb.Append('\n');
    }

    private sealed class NoBreakHyphenNode : IReplayNode
    {
        public void Replay(DxpIVisitor visitor, DxpIDocumentContext context) => visitor.VisitNoBreakHyphen(new NoBreakHyphen(), context);
        public void AppendText(StringBuilder sb) => sb.Append('-');
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
            if (node is HyperlinkNode)
                return false;
            if (node is RunNode runNode)
            {
                segments.Add((runNode.GetText(), runNode.CloneRunProperties()));
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
        foreach (var node in _nodes)
            node.AppendText(sb);
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
