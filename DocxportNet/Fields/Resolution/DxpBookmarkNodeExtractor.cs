using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet.Fields.Resolution;

internal static class DxpBookmarkNodeExtractor
{
    public static IReadOnlyDictionary<string, DxpFieldNodeBuffer> Extract(WordprocessingDocument document, ILogger? logger = null)
    {
        _ = logger;
        var captures = new Dictionary<string, BookmarkCapture>(StringComparer.OrdinalIgnoreCase);
        var idToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var activeIds = new List<string>();
        var part = document.MainDocumentPart;
        var body = part?.Document?.Body;
        if (body == null)
            return new Dictionary<string, DxpFieldNodeBuffer>(StringComparer.OrdinalIgnoreCase);

        void AddToActive(Action<BookmarkCapture> action)
        {
            if (activeIds.Count == 0)
                return;
            foreach (var id in activeIds)
            {
                if (!idToName.TryGetValue(id, out var name))
                    continue;
                if (!captures.TryGetValue(name, out var capture))
                    continue;
                action(capture);
            }
        }

        void Visit(OpenXmlElement element)
        {
            switch (element)
            {
                case BookmarkStart bs:
                {
                    string? name = bs.Name?.Value;
                    string? id = bs.Id?.Value;
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(id))
                        return;
                    idToName[id!] = name!;
                    activeIds.Add(id!);
                    if (!captures.ContainsKey(name!))
                        captures[name!] = new BookmarkCapture();
                    return;
                }
                case BookmarkEnd be:
                {
                    string? id = be.Id?.Value;
                    if (string.IsNullOrWhiteSpace(id))
                        return;
                    int idx = activeIds.LastIndexOf(id!);
                    if (idx >= 0)
                        activeIds.RemoveAt(idx);
                    return;
                }
                case Run run:
                {
                    if (activeIds.Count == 0)
                        return;
                    var cloned = DxpRunCloner.CloneRunWithParagraphAncestor(run);
                    var pops = new List<Action>(activeIds.Count);
                    AddToActive(capture => {
                        var child = capture.Current.BeginRun(cloned);
                        capture.Stack.Push(child);
                        pops.Add(() => capture.Stack.Pop());
                    });
                    foreach (var child in run.ChildElements)
                        Visit(child);
                    foreach (var pop in pops)
                        pop();
                    return;
                }
                case Hyperlink link:
                {
                    if (activeIds.Count == 0)
                        return;
                    var cloned = (Hyperlink)link.CloneNode(false);
                    var target = DxpHyperlinks.ResolveHyperlinkTarget(link, part);
                    var pops = new List<Action>(activeIds.Count);
                    AddToActive(capture => {
                        var child = capture.Current.BeginHyperlink(cloned, target);
                        capture.Stack.Push(child);
                        pops.Add(() => capture.Stack.Pop());
                    });
                    foreach (var child in link.ChildElements)
                        Visit(child);
                    foreach (var pop in pops)
                        pop();
                    return;
                }
                case Text text:
                    AddToActive(capture => capture.Current.AddText(text.Text));
                    return;
                case DeletedText dt:
                    AddToActive(capture => capture.Current.AddDeletedText(dt.Text));
                    return;
                case Break:
                    AddToActive(capture => capture.Current.AddBreak());
                    return;
                case TabChar:
                    AddToActive(capture => capture.Current.AddTab());
                    return;
                case CarriageReturn:
                    AddToActive(capture => capture.Current.AddCarriageReturn());
                    return;
                case NoBreakHyphen:
                    AddToActive(capture => capture.Current.AddNoBreakHyphen());
                    return;
            }

            if (element.HasChildren)
            {
                foreach (var child in element.ChildElements)
                    Visit(child);
            }
        }

        foreach (var child in body.ChildElements)
            Visit(child);

        var results = new Dictionary<string, DxpFieldNodeBuffer>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in captures)
            results[kvp.Key] = kvp.Value.Root;
        return results;
    }


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
}
