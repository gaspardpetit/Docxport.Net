using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Walker;
using Microsoft.Extensions.Logging;
using System.Runtime.CompilerServices;

namespace DocxportNet;

public static class DxpTransform
{
    public static void Transform(WordprocessingDocument document, IDxpNodeTransformer transformer, ILogger? logger = null)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        if (transformer == null)
            throw new ArgumentNullException(nameof(transformer));

        var engine = new DxpWordprocessingDocumentTransformEngine(logger);
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

internal sealed class DxpWordprocessingDocumentTransformEngine
{
    private readonly ILogger? _logger;

    public DxpWordprocessingDocumentTransformEngine(ILogger? logger)
    {
        _logger = logger;
    }

    public void Transform(WordprocessingDocument document, IDxpNodeTransformer transformer)
    {
        var enumerator = new DxpWordTransformPartEnumerator();
        var engine = new DxpOpenXmlTransformEngine(_logger, enumerator);
        engine.Transform(enumerator.Enumerate(document), transformer);
    }
}

internal interface IDxpTransformTraversalAdapter
{
    IEnumerable<OpenXmlElement> EnumerateChildren(OpenXmlElement node);
}

internal sealed record DxpTransformPartRoot(
    OpenXmlPart Part,
    DxpTransformPartKind PartKind,
    string RootName,
    IEnumerable<OpenXmlElement> RootNodes,
    Action Save);

internal sealed class DxpOpenXmlTransformEngine
{
    private readonly ILogger? _logger;
    private readonly IDxpTransformTraversalAdapter _adapter;

    public DxpOpenXmlTransformEngine(ILogger? logger, IDxpTransformTraversalAdapter adapter)
    {
        _logger = logger;
        _adapter = adapter;
    }

    public void Transform(IEnumerable<DxpTransformPartRoot> partRoots, IDxpNodeTransformer transformer)
    {
        foreach (var partRoot in partRoots)
        {
            int ordinal = 0;
            var ancestors = new List<OpenXmlElement>();
            var current = partRoot.RootNodes.FirstOrDefault();
            while (current != null)
            {
                var next = current.NextSibling<OpenXmlElement>();
                TransformNode(current, partRoot.Part, partRoot.PartKind, partRoot.RootName, partRoot.RootName, ancestors, ref ordinal, transformer);
                current = next;
            }

            partRoot.Save();
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
            foreach (var child in _adapter.EnumerateChildren(node).ToList())
                TransformNode(child, part, partKind, rootName, parentPath, ancestors, ref ordinal, transformer);
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

internal sealed class DxpWordTransformPartEnumerator : IDxpTransformTraversalAdapter
{
    private const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private readonly HashSet<OpenXmlElement> _seenFallbackNodes = new(ReferenceEqualityComparer.Instance);

    public IEnumerable<DxpTransformPartRoot> Enumerate(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("DOCX has no MainDocumentPart.");
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("DOCX has no main document body.");

        yield return new DxpTransformPartRoot(
            mainPart,
            DxpTransformPartKind.MainDocument,
            nameof(Body),
            body.ChildElements.OfType<OpenXmlElement>(),
            () => mainPart.Document!.Save());

        foreach (var entry in EnumerateReachableHeaderFooterParts(body, mainPart))
        {
            if (entry.root == null)
                continue;

            yield return new DxpTransformPartRoot(
                entry.part,
                entry.kind,
                entry.root.LocalName,
                entry.root.ChildElements.OfType<OpenXmlElement>(),
                entry.root.Save);
        }

        if (mainPart.FootnotesPart?.Footnotes != null)
        {
            yield return new DxpTransformPartRoot(
                mainPart.FootnotesPart,
                DxpTransformPartKind.Footnote,
                mainPart.FootnotesPart.Footnotes.LocalName,
                mainPart.FootnotesPart.Footnotes.Elements<Footnote>()
                    .Where(static fn => !IsInternalNote(fn.Type?.Value))
                    .Cast<OpenXmlElement>(),
                () => mainPart.FootnotesPart.Footnotes.Save());
        }

        if (mainPart.EndnotesPart?.Endnotes != null)
        {
            yield return new DxpTransformPartRoot(
                mainPart.EndnotesPart,
                DxpTransformPartKind.Endnote,
                mainPart.EndnotesPart.Endnotes.LocalName,
                mainPart.EndnotesPart.Endnotes.Elements<Endnote>()
                    .Where(static en => !IsInternalNote(en.Type?.Value))
                    .Cast<OpenXmlElement>(),
                () => mainPart.EndnotesPart.Endnotes.Save());
        }
    }

    public IEnumerable<OpenXmlElement> EnumerateChildren(OpenXmlElement node)
    {
        if (IsInsideOpaqueWordprocessingTextBoxSubtree(node))
            yield break;

        var current = node.FirstChild;
        while (current != null)
        {
            var next = current.NextSibling();
            if (current is OpenXmlElement child)
                yield return child;
            current = next;
        }

        if (IsWordprocessingTextBoxContent(node))
            yield break;

        if (TryEnumerateTextBoxCarrierChildren(node, out var carrierChildren))
        {
            foreach (var child in carrierChildren)
            {
                if (_seenFallbackNodes.Add(child) && !IsAlreadyReachableFromNormalChildren(node, child))
                    yield return child;
            }
        }
    }

    private static bool IsInternalNote(FootnoteEndnoteValues? type)
        => type == FootnoteEndnoteValues.Separator
            || type == FootnoteEndnoteValues.ContinuationSeparator
            || type == FootnoteEndnoteValues.ContinuationNotice;

    private static IEnumerable<(OpenXmlPart part, OpenXmlPartRootElement? root, DxpTransformPartKind kind)> EnumerateReachableHeaderFooterParts(Body body, MainDocumentPart mainPart)
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

    private static bool IsAlreadyReachableFromNormalChildren(OpenXmlElement node, OpenXmlElement candidate)
    {
        foreach (var child in node.ChildElements.OfType<OpenXmlElement>())
        {
            if (ReferenceEquals(child, candidate))
                return true;

            if (child.Descendants().Any(descendant => ReferenceEquals(descendant, candidate)))
                return true;
        }

        return false;
    }

    private static bool TryEnumerateTextBoxCarrierChildren(OpenXmlElement node, out IReadOnlyList<OpenXmlElement> children)
    {
        if (node is Drawing drawing)
        {
            children = EnumerateTextBoxContentChildren(
                drawing.Descendants<TextBoxInfo2>()
                    .Select(static txbx => txbx.GetFirstChild<TextBoxContent>())
                    .Where(static content => content != null)
                    .Select(static content => content!)).ToList();
            return children.Count > 0;
        }

        if (node is Picture picture)
        {
            children = EnumerateTextBoxContentChildren(
                picture.Descendants<DocumentFormat.OpenXml.Vml.TextBox>()
                    .Select(static txbx => txbx.GetFirstChild<TextBoxContent>())
                    .Where(static content => content != null)
                    .Select(static content => content!)).ToList();
            return children.Count > 0;
        }

        if (IsTextBoxCarrier(node) && TryGetNestedWordprocessingTextBoxContent(node, out var content))
        {
            children = EnumerateElementChildren(content).ToList();
            return children.Count > 0;
        }

        children = Array.Empty<OpenXmlElement>();
        return false;
    }

    private static IEnumerable<OpenXmlElement> EnumerateTextBoxContentChildren(IEnumerable<TextBoxContent> contents)
    {
        foreach (var content in contents)
        {
            var current = content.FirstChild;
            while (current != null)
            {
                var next = current.NextSibling();
                if (current is OpenXmlElement child)
                    yield return child;
                current = next;
            }
        }
    }

    private static IEnumerable<OpenXmlElement> EnumerateElementChildren(OpenXmlElement node)
    {
        var current = node.FirstChild;
        while (current != null)
        {
            var next = current.NextSibling();
            if (current is OpenXmlElement child)
                yield return child;
            current = next;
        }
    }

    private static bool TryGetNestedWordprocessingTextBoxContent(OpenXmlElement node, out OpenXmlElement content)
    {
        foreach (var descendant in node.Descendants())
        {
            if (!IsWordprocessingTextBoxContent(descendant))
                continue;

            content = descendant;
            return true;
        }

        content = null!;
        return false;
    }

    private static bool IsWordprocessingTextBoxContent(OpenXmlElement node)
        => string.Equals(node.LocalName, "txbxContent", StringComparison.Ordinal)
            && string.Equals(node.NamespaceUri, WordprocessingNamespace, StringComparison.Ordinal);

    private static bool IsTextBoxCarrier(OpenXmlElement node)
        => node is Picture
            || node is Drawing
            || (string.Equals(node.LocalName, "shape", StringComparison.Ordinal)
                && string.Equals(node.NamespaceUri, "urn:schemas-microsoft-com:vml", StringComparison.Ordinal))
            || (string.Equals(node.LocalName, "txbx", StringComparison.Ordinal)
                && node.NamespaceUri.Contains("wordprocessingShape"));

    private static bool IsInsideOpaqueWordprocessingTextBoxSubtree(OpenXmlElement node)
    {
        if (IsWordprocessingTextBoxContent(node))
            return false;

        var parent = node.Parent;
        while (parent != null)
        {
            if (IsWordprocessingTextBoxContent(parent))
                return true;

            parent = parent.Parent;
        }

        return false;
    }

    private sealed class ReferenceEqualityComparer : IEqualityComparer<OpenXmlElement>
    {
        public static readonly ReferenceEqualityComparer Instance = new();

        public bool Equals(OpenXmlElement? x, OpenXmlElement? y) => ReferenceEquals(x, y);

        public int GetHashCode(OpenXmlElement obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
