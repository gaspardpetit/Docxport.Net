namespace DocxportNet.API;

public sealed class DxpRefIndex
{
    public IReadOnlyDictionary<string, DxpRefBookmark> Bookmarks { get; }
    public IReadOnlyDictionary<string, DxpRefFootnote> Footnotes { get; }
    public IReadOnlyDictionary<string, DxpRefEndnote> Endnotes { get; }
    public IReadOnlyDictionary<string, DxpRefHyperlink> Hyperlinks { get; }
    public IReadOnlyDictionary<string, DxpRefParagraphNumber> ParagraphNumbers { get; }

    public DxpRefIndex(
        IReadOnlyDictionary<string, DxpRefBookmark> bookmarks,
        IReadOnlyDictionary<string, DxpRefFootnote> footnotes,
        IReadOnlyDictionary<string, DxpRefEndnote> endnotes,
        IReadOnlyDictionary<string, DxpRefHyperlink> hyperlinks,
        IReadOnlyDictionary<string, DxpRefParagraphNumber> paragraphNumbers)
    {
        Bookmarks = bookmarks;
        Footnotes = footnotes;
        Endnotes = endnotes;
        Hyperlinks = hyperlinks;
        ParagraphNumbers = paragraphNumbers;
    }

    public static DxpRefIndex Empty { get; } = new DxpRefIndex(
        new Dictionary<string, DxpRefBookmark>(),
        new Dictionary<string, DxpRefFootnote>(),
        new Dictionary<string, DxpRefEndnote>(),
        new Dictionary<string, DxpRefHyperlink>(),
        new Dictionary<string, DxpRefParagraphNumber>());
}

public sealed record DxpRefBookmark(
    string Name,
    string Text,
    int DocumentOrder
);

public sealed record DxpRefFootnote(
    string Bookmark,
    string Mark,
    string Text
);

public sealed record DxpRefEndnote(
    string Bookmark,
    string Mark,
    string Text
);

public sealed record DxpRefHyperlink(
    string Bookmark,
    string Target,
    string? Text
);

public sealed record DxpRefParagraphNumber(
    string Bookmark,
    string FullNumber,
    string CurrentLevelNumber,
    string NumericOnly
);
