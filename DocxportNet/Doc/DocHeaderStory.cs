using System.Globalization;

namespace DocxportNet.Doc;

public enum DocHeaderStoryKind
{
    FootnoteSeparator, FootnoteContinuationSeparator, FootnoteContinuationNotice,
    EndnoteSeparator, EndnoteContinuationSeparator, EndnoteContinuationNotice,
    EvenHeader, OddHeader, EvenFooter, OddFooter, FirstHeader, FirstFooter
}

/// <summary>A Plcfhdd slot in the Header Document, addressed in global CP space.</summary>
public sealed class DocHeaderStory
{
    private readonly DocTextIndex _index;

    internal DocHeaderStory(DocTextIndex index, DocStructureNode node, int slot)
    {
        _index = index;
        Slot = slot;
        CpStart = uint.Parse(node.Attributes["globalCpStart"], CultureInfo.InvariantCulture);
        CpEnd = uint.Parse(node.Attributes["globalCpEnd"], CultureInfo.InvariantCulture);
        SectionIndex = slot < 6 ? null : (slot - 6) / 6;
        Kind = slot < 6 ? (DocHeaderStoryKind)slot :
            (DocHeaderStoryKind)(6 + (slot - 6) % 6);
    }

    public int Slot { get; }
    public int? SectionIndex { get; }
    public DocHeaderStoryKind Kind { get; }
    public uint CpStart { get; }
    public uint CpEnd { get; }
    public bool IsEmpty => CpStart == CpEnd;

    /// <summary>Reads visible content with formatting scoped to this slot.</summary>
    public DocIndexedStoryText ReadIndexedContent() => _index.BindStory(ReadContent());

    /// <summary>Reads content without the mandatory final guard paragraph mark.</summary>
    public DocStoryText ReadContent()
    {
        if (IsEmpty) return DocStoryTextReader.ReadRange(_index, "Headers", CpStart, CpEnd);
        var guard = DocStoryTextReader.ReadRange(_index, "Headers", CpEnd - 1, CpEnd);
        if (guard.Paragraphs.Count != 1 || guard.Paragraphs[0].End != DocParagraphEnd.ParagraphMark)
            throw new InvalidDataException($"Header story {Slot} has no guard paragraph mark.");
        return DocStoryTextReader.ReadRange(_index, "Headers", CpStart, CpEnd - 1);
    }
}
