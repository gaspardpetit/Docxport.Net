using System.Globalization;
using DocxportNet.Core;

namespace DocxportNet.Doc;

/// <summary>Indexes logical text positions without decoding the referenced text bytes.</summary>
public sealed class DocTextIndexWalker
{
    public DocTextIndex Index(string path)
    {
        if (path == null) throw new ArgumentNullException(nameof(path));
        var collector = new PieceCollector();
        return Build(new DocStructureWalker().Accept(path, collector), collector.Pieces);
    }

    public DocTextIndex Index(Stream input)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        var collector = new PieceCollector();
        return Build(new DocStructureWalker().Accept(input, collector), collector.Pieces);
    }

    private static DocTextIndex Build(DocStructure structure, IReadOnlyList<DocStructureNode> nodes)
    {
        try
        {
            var pieces = nodes.Select(x => new DocIndexedTextPiece(x)).ToArray();
            if (pieces.Length == 0 && structure.Parts.Count != 0)
                throw new InvalidDataException("The DOC has document text but no text piece table.");
            return new DocTextIndex(structure, pieces);
        }
        catch
        {
            structure.Dispose();
            throw;
        }
    }

    private sealed class PieceCollector : IDocStructureVisitor
    {
        public List<DocStructureNode> Pieces { get; } = new();

        public IDisposable? Enter(DocStructureNode node, int depth)
        {
            if (node.Kind == "Pcd") Pieces.Add(node);
            return node.Kind switch
            {
                "Document" or "FIB" or "FibDirectory" or "Pcdt" or "PlcPcd" => DxpDisposable.Empty,
                "Stream" when node.Name == "WordDocument" => DxpDisposable.Empty,
                "FibLocation" when node.Name == "TextPieceTable" => DxpDisposable.Empty,
                _ => null
            };
        }
    }
}

/// <summary>A referenced formatting page; run headers are decoded on first access.</summary>
public sealed class DocIndexedFormattingPage
{
    private readonly DocStructure _structure;

    internal DocIndexedFormattingPage(DocStructure structure, DocStructureNode node)
    {
        _structure = structure;
        Node = node;
        FcStart = uint.Parse(node.Attributes["fcStart"], CultureInfo.InvariantCulture);
        FcEnd = uint.Parse(node.Attributes["fcEnd"], CultureInfo.InvariantCulture);
        PageNumber = uint.Parse(node.Attributes["pageNumber"], CultureInfo.InvariantCulture);
    }

    public DocStructureNode Node { get; }
    public bool IsCharacterFormatting => Node.Kind == "ChpxFkp";
    public uint FcStart { get; }
    public uint FcEnd { get; }
    public uint PageNumber { get; }
    public bool AreRunsLoaded => Node.Children.Count != 0;
    public IReadOnlyList<DocStructureNode> Runs
    {
        get
        {
            DocFormattingNavigator.ExpandPage(_structure, Node);
            return (IReadOnlyList<DocStructureNode>)Node.Children;
        }
    }
}

/// <summary>An ordered logical text piece backed by a Pcd and a WordDocument byte range.</summary>
public sealed class DocIndexedTextPiece
{
    internal DocIndexedTextPiece(DocStructureNode node)
    {
        Node = node;
        CpStart = uint.Parse(node.Attributes["cpStart"], CultureInfo.InvariantCulture);
        CpEnd = uint.Parse(node.Attributes["cpEnd"], CultureInfo.InvariantCulture);
        TextOffset = long.Parse(node.Attributes["textOffset"], CultureInfo.InvariantCulture);
        TextByteLength = long.Parse(node.Attributes["textLength"], CultureInfo.InvariantCulture);
        Encoding = node.Attributes["encoding"];
    }

    public DocStructureNode Node { get; }
    public uint CpStart { get; }
    public uint CpEnd { get; }
    public long TextOffset { get; }
    public long TextByteLength { get; }
    public string Encoding { get; }
    public bool IsTextLoaded => Node.IsPayloadLoaded;
    public string Text => ((DocTextPieceContent)(Node.Payload ??
        throw new InvalidDataException("The text piece has no payload."))).Text;
}

/// <summary>A document-part slice of one text piece. Decoding is deferred until Text is read.</summary>
public sealed class DocIndexedTextSpan
{
    internal DocIndexedTextSpan(DocIndexedTextPiece piece, uint cpStart, uint cpEnd)
    {
        Piece = piece;
        CpStart = cpStart;
        CpEnd = cpEnd;
    }

    public DocIndexedTextPiece Piece { get; }
    public uint CpStart { get; }
    public uint CpEnd { get; }
    public string Text => Piece.Text.Substring(checked((int)(CpStart - Piece.CpStart)),
        checked((int)(CpEnd - CpStart)));
}

/// <summary>Owns the source DOC while indexed text ranges may be materialized on demand.</summary>
public sealed class DocTextIndex : IDisposable
{
    private readonly Lazy<IReadOnlyList<DocIndexedFormattingPage>> _formattingPages;
    private readonly Lazy<IReadOnlyList<DocStructureNode>> _styles;
    private readonly Lazy<IReadOnlyList<DocStructureNode>> _sections;

    internal DocTextIndex(DocStructure structure, IReadOnlyList<DocIndexedTextPiece> pieces)
    {
        Structure = structure;
        Pieces = pieces;
        _formattingPages = new Lazy<IReadOnlyList<DocIndexedFormattingPage>>(LoadFormattingPages);
        _styles = new Lazy<IReadOnlyList<DocStructureNode>>(LoadStyles);
        _sections = new Lazy<IReadOnlyList<DocStructureNode>>(LoadSections);
    }

    public DocStructure Structure { get; }
    public IReadOnlyList<DocIndexedTextPiece> Pieces { get; }
    public IReadOnlyList<DocIndexedFormattingPage> FormattingPages => _formattingPages.Value;
    public IReadOnlyList<DocStructureNode> Styles => _styles.Value;
    public IReadOnlyList<DocStructureNode> Sections => _sections.Value;
    public bool IsFormattingIndexLoaded => _formattingPages.IsValueCreated;
    public bool IsStyleIndexLoaded => _styles.IsValueCreated;
    public bool IsSectionIndexLoaded => _sections.IsValueCreated;
    public IReadOnlyList<DocPartRange> Parts => Structure.Parts;

    /// <summary>Returns logical-order slices for a document part without reading text bytes.</summary>
    public IReadOnlyList<DocIndexedTextSpan> GetPartSpans(string partName)
    {
        if (partName == null) throw new ArgumentNullException(nameof(partName));
        var part = Parts.FirstOrDefault(x => x.Name == partName);
        if (part == null) return Array.Empty<DocIndexedTextSpan>();
        var spans = new List<DocIndexedTextSpan>();
        foreach (var piece in Pieces)
        {
            var start = Math.Max(piece.CpStart, part.CpStart);
            var end = Math.Min(piece.CpEnd, part.CpEnd);
            if (start < end) spans.Add(new DocIndexedTextSpan(piece, start, end));
        }
        return spans;
    }

    private IReadOnlyList<DocIndexedFormattingPage> LoadFormattingPages()
    {
        var pages = new List<DocIndexedFormattingPage>();
        foreach (var name in new[] { "CharacterFormatting", "ParagraphFormatting" })
        {
            var location = FindLocationNode(name);
            if (location == null) continue;
            DocFormattingNavigator.ExpandPageTable(Structure, location);
            foreach (var page in location.Children.SelectMany(plc => plc.Children)
                         .SelectMany(bte => bte.Children))
                pages.Add(new DocIndexedFormattingPage(Structure, page));
        }
        return pages;
    }

    private IReadOnlyList<DocStructureNode> LoadStyles()
    {
        var location = FindLocationNode("StyleSheet");
        if (location == null) return Array.Empty<DocStructureNode>();
        DocStyleSheetNavigator.Expand(Structure, location);
        return location.Children.SelectMany(stsh => stsh.Children)
            .Where(node => node.Kind == "LPStd").ToArray();
    }

    private IReadOnlyList<DocStructureNode> LoadSections()
    {
        var location = FindLocationNode("Sections");
        if (location == null) return Array.Empty<DocStructureNode>();
        DocSectionNavigator.Expand(Structure, location);
        return location.Children.SelectMany(plc => plc.Children).ToArray();
    }

    private DocStructureNode? FindLocationNode(string name) => Structure.Root.Children
        .Where(node => node.Name == "WordDocument")
        .SelectMany(node => node.Children)
        .Where(node => node.Kind == "FIB")
        .SelectMany(node => node.Children)
        .Where(node => node.Kind == "FibDirectory")
        .SelectMany(node => node.Children)
        .FirstOrDefault(node => node.Name == name);

    public void Dispose() => Structure.Dispose();
}
