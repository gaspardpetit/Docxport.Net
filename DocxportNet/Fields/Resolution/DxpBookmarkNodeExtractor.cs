using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Core;
using DocxportNet.Visitors;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Resolution;

internal static class DxpBookmarkNodeExtractor
{
    public static IReadOnlyDictionary<string, DxpFieldNodeBuffer> Extract(WordprocessingDocument document, ILogger? logger = null)
    {
        var visitor = new DxpBookmarkNodeVisitor(logger);
        var pipeline = DxpVisitorMiddleware.Chain(
            visitor,
            next => new DxpContextTracker(next));
        var walker = new DxpWalker(logger);
        walker.Accept(document, pipeline);
        return visitor.Results;
    }

    private sealed class DxpBookmarkNodeVisitor : DxpVisitor
    {
        private sealed class BookmarkCapture
        {
            public DxpFieldNodeBuffer Root { get; } = new();
            public Stack<DxpFieldNodeBuffer> Stack { get; } = new();

            public BookmarkCapture()
            {
                Stack.Push(Root);
            }

            public DxpFieldNodeBuffer Current => Stack.Peek();
        }

        private readonly Dictionary<string, BookmarkCapture> _captures = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _idToName = new();
        private readonly List<string> _activeIds = new();

        public DxpBookmarkNodeVisitor(ILogger? logger) : base(logger)
        {
        }

        public IReadOnlyDictionary<string, DxpFieldNodeBuffer> Results
        {
            get
            {
                var results = new Dictionary<string, DxpFieldNodeBuffer>(StringComparer.OrdinalIgnoreCase);
                foreach (var kvp in _captures)
                    results[kvp.Key] = kvp.Value.Root;
                return results;
            }
        }

        public override void VisitBookmarkStart(BookmarkStart bs, DxpIDocumentContext d)
        {
            string? name = bs.Name?.Value;
            string? id = bs.Id?.Value;
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
                return;

            _idToName[id!] = name!;
            _activeIds.Add(id!);
            if (!_captures.ContainsKey(name!))
                _captures[name!] = new BookmarkCapture();
        }

        public override void VisitBookmarkEnd(BookmarkEnd be, DxpIDocumentContext d)
        {
            string? id = be.Id?.Value;
            if (string.IsNullOrWhiteSpace(id))
                return;

            int idx = _activeIds.LastIndexOf(id!);
            if (idx >= 0)
                _activeIds.RemoveAt(idx);
        }

        public override IDisposable VisitRunBegin(Run r, DxpIDocumentContext d)
        {
            if (_activeIds.Count == 0)
                return DxpDisposable.Empty;

            var run = new Run();
            if (r.RunProperties != null)
                run.Append(r.RunProperties.CloneNode(true));
            var scopes = new List<Action>(_activeIds.Count);
            foreach (var id in _activeIds)
            {
                if (!_idToName.TryGetValue(id, out var name))
                    continue;
                if (!_captures.TryGetValue(name, out var capture))
                    continue;
                var child = capture.Current.BeginRun(run);
                capture.Stack.Push(child);
                scopes.Add(() => capture.Stack.Pop());
            }

            return DxpDisposable.Create(() => {
                foreach (var pop in scopes)
                    pop();
            });
        }

        public override IDisposable VisitHyperlinkBegin(Hyperlink link, DxpLinkAnchor? target, DxpIDocumentContext d)
        {
            if (_activeIds.Count == 0)
                return DxpDisposable.Empty;

            var cloned = (Hyperlink)link.CloneNode(false);
            var scopes = new List<Action>(_activeIds.Count);
            foreach (var id in _activeIds)
            {
                if (!_idToName.TryGetValue(id, out var name))
                    continue;
                if (!_captures.TryGetValue(name, out var capture))
                    continue;
                var child = capture.Current.BeginHyperlink(cloned, target);
                capture.Stack.Push(child);
                scopes.Add(() => capture.Stack.Pop());
            }

            return DxpDisposable.Create(() => {
                foreach (var pop in scopes)
                    pop();
            });
        }

        public override void VisitText(Text t, DxpIDocumentContext d) => AppendText(t.Text);
        public override void VisitDeletedText(DeletedText dt, DxpIDocumentContext d) => AppendDeletedText(dt.Text);
        public override void VisitTab(TabChar tab, DxpIDocumentContext d) => AppendTab();
        public override void VisitBreak(Break br, DxpIDocumentContext d) => AppendBreak();
        public override void VisitCarriageReturn(CarriageReturn cr, DxpIDocumentContext d) => AppendCarriageReturn();
        public override void VisitNoBreakHyphen(NoBreakHyphen h, DxpIDocumentContext d) => AppendNoBreakHyphen();

        private void AppendText(string text)
        {
            foreach (var capture in EnumerateActiveCaptures())
                capture.Current.AddText(text);
        }

        private void AppendDeletedText(string text)
        {
            foreach (var capture in EnumerateActiveCaptures())
                capture.Current.AddDeletedText(text);
        }

        private void AppendTab()
        {
            foreach (var capture in EnumerateActiveCaptures())
                capture.Current.AddTab();
        }

        private void AppendBreak()
        {
            foreach (var capture in EnumerateActiveCaptures())
                capture.Current.AddBreak();
        }

        private void AppendCarriageReturn()
        {
            foreach (var capture in EnumerateActiveCaptures())
                capture.Current.AddCarriageReturn();
        }

        private void AppendNoBreakHyphen()
        {
            foreach (var capture in EnumerateActiveCaptures())
                capture.Current.AddNoBreakHyphen();
        }

        private IEnumerable<BookmarkCapture> EnumerateActiveCaptures()
        {
            foreach (var id in _activeIds)
            {
                if (_idToName.TryGetValue(id, out var name) && _captures.TryGetValue(name, out var capture))
                    yield return capture;
            }
        }
    }
}
