using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;

namespace DocxportNet;

public static class DxpTransform
{
    public static void Transform(WordprocessingDocument document, IDxpNodeTransformer transformer, ILogger? logger = null)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        if (transformer == null)
            throw new ArgumentNullException(nameof(transformer));

        var engine = new DxpDocumentTransformEngine(logger);
        engine.Transform(document, transformer);
    }

    public static void Transform(string inputPath, string outputPath, IDxpNodeTransformer transformer, ILogger? logger = null)
    {
        if (string.IsNullOrWhiteSpace(inputPath))
            throw new ArgumentException("Input path is required.", nameof(inputPath));
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new ArgumentException("Output path is required.", nameof(outputPath));
        if (transformer == null)
            throw new ArgumentNullException(nameof(transformer));

        string? directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        File.Copy(inputPath, outputPath, overwrite: true);
        using var document = WordprocessingDocument.Open(outputPath, true);
        Transform(document, transformer, logger);
    }

    public static byte[] Transform(byte[] docxBytes, IDxpNodeTransformer transformer, ILogger? logger = null)
    {
        if (docxBytes == null)
            throw new ArgumentNullException(nameof(docxBytes));
        if (transformer == null)
            throw new ArgumentNullException(nameof(transformer));

        using var stream = new MemoryStream();
        stream.Write(docxBytes, 0, docxBytes.Length);
        stream.Position = 0;

        using (var document = WordprocessingDocument.Open(stream, true))
            Transform(document, transformer, logger);

        return stream.ToArray();
    }
}

public enum DxpTransformPartKind
{
    MainDocument,
    Header,
    Footer,
    Footnote,
    Endnote
}

public enum DxpTransformAction
{
    Keep,
    Remove,
    Replace
}

public sealed class DxpTransformDecision
{
    private static readonly IReadOnlyList<OpenXmlElement> EmptyReplacements = Array.Empty<OpenXmlElement>();

    private DxpTransformDecision(DxpTransformAction action, bool descend, IReadOnlyList<OpenXmlElement> replacements)
    {
        Action = action;
        Descend = descend;
        Replacements = replacements;
    }

    public DxpTransformAction Action { get; }
    public bool Descend { get; }
    public IReadOnlyList<OpenXmlElement> Replacements { get; }

    public static DxpTransformDecision Keep(bool descend = true)
        => new(DxpTransformAction.Keep, descend, EmptyReplacements);

    public static DxpTransformDecision Remove()
        => new(DxpTransformAction.Remove, descend: false, EmptyReplacements);

    public static DxpTransformDecision Replace(params OpenXmlElement[] replacements)
        => Replace((IEnumerable<OpenXmlElement>)replacements);

    public static DxpTransformDecision Replace(IEnumerable<OpenXmlElement> replacements)
    {
        if (replacements == null)
            throw new ArgumentNullException(nameof(replacements));

        return new DxpTransformDecision(
            DxpTransformAction.Replace,
            descend: false,
            replacements.Where(static r => r != null)
                .Select(static r => r!)
                .ToArray());
    }
}

public sealed class DxpTransformContext
{
    internal DxpTransformContext(
        OpenXmlElement node,
        OpenXmlPart part,
        DxpTransformPartKind partKind,
        IReadOnlyList<OpenXmlElement> ancestors,
        int depth,
        int siblingIndex,
        int nodeOrdinal,
        string path)
    {
        Node = node;
        Part = part;
        PartKind = partKind;
        Ancestors = ancestors;
        Depth = depth;
        SiblingIndex = siblingIndex;
        NodeOrdinal = nodeOrdinal;
        Path = path;
    }

    public OpenXmlElement Node { get; }
    public OpenXmlPart Part { get; }
    public DxpTransformPartKind PartKind { get; }
    public IReadOnlyList<OpenXmlElement> Ancestors { get; }
    public int Depth { get; }
    public int SiblingIndex { get; }
    public int NodeOrdinal { get; }
    public string Path { get; }
}

public interface IDxpNodeTransformer
{
    DxpTransformDecision Visit(OpenXmlElement node, DxpTransformContext context);
}

internal sealed class DxpDocumentTransformEngine
{
    private readonly ILogger? _logger;

    public DxpDocumentTransformEngine(ILogger? logger)
    {
        _logger = logger;
    }

    public void Transform(WordprocessingDocument document, IDxpNodeTransformer transformer)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("DOCX has no MainDocumentPart.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("DOCX has no main document body.");

        TransformPart(
            body.ChildElements.OfType<OpenXmlElement>(),
            mainPart,
            DxpTransformPartKind.MainDocument,
            nameof(Body));
        mainPart.Document!.Save();

        foreach (var entry in EnumerateReachableHeaderFooterParts(body, mainPart))
        {
            if (entry.root == null)
                continue;

            TransformPart(
                entry.root.ChildElements.OfType<OpenXmlElement>(),
                entry.part,
                entry.kind,
                entry.root.LocalName);
            entry.root.Save();
        }

        if (mainPart.FootnotesPart?.Footnotes != null)
        {
            TransformPart(
                mainPart.FootnotesPart.Footnotes.Elements<Footnote>()
                    .Where(static fn => !IsInternalNote(fn.Type?.Value))
                    .Cast<OpenXmlElement>(),
                mainPart.FootnotesPart,
                DxpTransformPartKind.Footnote,
                mainPart.FootnotesPart.Footnotes.LocalName);
            mainPart.FootnotesPart.Footnotes.Save();
        }

        if (mainPart.EndnotesPart?.Endnotes != null)
        {
            TransformPart(
                mainPart.EndnotesPart.Endnotes.Elements<Endnote>()
                    .Where(static en => !IsInternalNote(en.Type?.Value))
                    .Cast<OpenXmlElement>(),
                mainPart.EndnotesPart,
                DxpTransformPartKind.Endnote,
                mainPart.EndnotesPart.Endnotes.LocalName);
            mainPart.EndnotesPart.Endnotes.Save();
        }

        void TransformPart(
            IEnumerable<OpenXmlElement> rootNodes,
            OpenXmlPart part,
            DxpTransformPartKind partKind,
            string rootName)
        {
            int ordinal = 0;
            var ancestors = new List<OpenXmlElement>();
            var current = rootNodes.FirstOrDefault();
            while (current != null)
            {
                var next = current.NextSibling<OpenXmlElement>();
                TransformNode(current, part, partKind, rootName, parentPath: rootName, ancestors, ref ordinal, transformer);
                current = next;
            }
        }
    }

    private static bool IsInternalNote(FootnoteEndnoteValues? type)
        => type == FootnoteEndnoteValues.Separator
            || type == FootnoteEndnoteValues.ContinuationSeparator
            || type == FootnoteEndnoteValues.ContinuationNotice;

    private IEnumerable<(OpenXmlPart part, OpenXmlPartRootElement? root, DxpTransformPartKind kind)> EnumerateReachableHeaderFooterParts(Body body, MainDocumentPart mainPart)
    {
        var seen = new HashSet<Uri>();
        foreach (var section in DxpSections.SplitDocumentBodyIntoSections(body))
        {
            foreach (var child in section.Properties.ChildElements)
            {
                switch (child)
                {
                    case HeaderReference headerRef when headerRef.Id?.Value != null:
                    {
                        if (mainPart.GetPartById(headerRef.Id.Value) is HeaderPart headerPart && seen.Add(headerPart.Uri))
                            yield return (headerPart, headerPart.Header, DxpTransformPartKind.Header);
                        break;
                    }
                    case FooterReference footerRef when footerRef.Id?.Value != null:
                    {
                        if (mainPart.GetPartById(footerRef.Id.Value) is FooterPart footerPart && seen.Add(footerPart.Uri))
                            yield return (footerPart, footerPart.Footer, DxpTransformPartKind.Footer);
                        break;
                    }
                }
            }
        }
    }

    private void TransformNode(
        OpenXmlElement node,
        OpenXmlPart part,
        DxpTransformPartKind partKind,
        string rootName,
        string parentPath,
        List<OpenXmlElement> ancestors,
        ref int ordinal,
        IDxpNodeTransformer transformer)
    {
        var parent = node.Parent;
        if (parent == null)
            return;

        int siblingIndex = GetSiblingIndex(parent, node);
        int depth = ancestors.Count;
        string path = string.Concat(parentPath, "/", GetPathSegment(node, siblingIndex));
        ordinal++;

        var context = new DxpTransformContext(
            node,
            part,
            partKind,
            ancestors.ToArray(),
            depth,
            siblingIndex,
            ordinal,
            path);

        var decision = transformer.Visit(node, context) ?? DxpTransformDecision.Keep();
        _logger?.LogDebug(
            "Transform visit: Part={PartKind} Path={Path} Action={Action} Descend={Descend}",
            partKind,
            path,
            decision.Action,
            decision.Descend);

        switch (decision.Action)
        {
            case DxpTransformAction.Keep:
                if (decision.Descend)
                    TransformChildren(node, part, partKind, rootName, path, ancestors, ref ordinal, transformer);
                break;

            case DxpTransformAction.Remove:
                node.Remove();
                break;

            case DxpTransformAction.Replace:
                ReplaceNode(node, decision.Replacements);
                break;
        }
    }

    private void TransformChildren(
        OpenXmlElement node,
        OpenXmlPart part,
        DxpTransformPartKind partKind,
        string rootName,
        string parentPath,
        List<OpenXmlElement> ancestors,
        ref int ordinal,
        IDxpNodeTransformer transformer)
    {
        ancestors.Add(node);
        try
        {
            var current = node.FirstChild;
            while (current != null)
            {
                var next = current.NextSibling();
                if (current is OpenXmlElement child)
                    TransformNode(child, part, partKind, rootName, parentPath, ancestors, ref ordinal, transformer);
                current = next;
            }
        }
        finally
        {
            ancestors.RemoveAt(ancestors.Count - 1);
        }
    }

    private static void ReplaceNode(OpenXmlElement node, IReadOnlyList<OpenXmlElement> replacements)
    {
        var parent = node.Parent;
        if (parent == null)
            return;

        OpenXmlElement anchor = node;
        foreach (var replacement in replacements)
        {
            var clone = (OpenXmlElement)replacement.CloneNode(true);
            anchor = parent.InsertAfter(clone, anchor)!;
        }

        node.Remove();
    }

    private static string GetPathSegment(OpenXmlElement node, int siblingIndex)
        => string.Concat(GetNodeName(node), "[", siblingIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), "]");

    private static string GetNodeName(OpenXmlElement node)
        => node.GetType().Name;

    private static int GetSiblingIndex(OpenXmlElement parent, OpenXmlElement node)
    {
        int index = 0;
        foreach (var child in parent.ChildElements)
        {
            if (ReferenceEquals(child, node))
                return index;
            index++;
        }

        return -1;
    }
}
