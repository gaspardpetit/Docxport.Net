namespace DocxportNet.Doc;

/// <summary>Section stories in DOC header-table order: even/odd headers,
/// even/odd footers, first header/footer. Null means inherit or absent.</summary>
internal sealed record DocPlainTextStory(string Text,
    IReadOnlyList<DocStoryCharacterRun> Runs,
    IReadOnlyList<DocStoryParagraphStyleRun> ParagraphStyles)
{
    public IReadOnlyList<DocStoryParagraphRange> Paragraphs { get; init; } = [];
    public IReadOnlyList<DocStoryTableRow> TableRows { get; init; } = [];
    public IReadOnlyList<DocStoryListParagraph>? ListParagraphs { get; init; }
    public IReadOnlyList<DocStoryFieldMark>? FieldMarks { get; init; }
    public IReadOnlyList<DocStoryBookmark> Bookmarks { get; init; } = [];
    public IReadOnlyList<DocStoryInlinePicture> Pictures { get; init; } = [];
    public IReadOnlyList<DocStoryFloatingPicture> FloatingPictures { get; init; } = [];
}

internal sealed record DocPlainTextSection(int EndCp,
    IReadOnlyList<DocPlainTextStory?> Stories,
    DocSectionFormatting? Formatting = null)
{
    public DocStorySection? Model { get; init; }
}

internal sealed record DocPlainTextDocument(DocPlainTextStory MainStory,
    IReadOnlyList<DocPlainTextSection> Sections,
    IReadOnlyList<DocStyleDefinition>? Styles = null,
    DocCharacterFormatting? DefaultCharacterFormatting = null,
    bool EvenAndOddHeaders = false,
    DocListIndex? Lists = null,
    DocParagraphFormatting? DefaultParagraphFormatting = null,
    short DefaultTabStopTwips = 720,
    bool MirrorMargins = false,
    bool GutterAtTop = false,
    IReadOnlyList<DocFontDefinition>? FontMetadata = null,
    bool BalanceSingleByteDoubleByteWidth = false,
    bool AutoHyphenation = false,
    short HyphenationZoneTwips = 0,
    short ConsecutiveHyphenLimit = 0,
    bool HyphenateCaps = true,
    string? Title = null,
    string? Subject = null,
    string? Author = null,
    string? Keywords = null,
    string? Comments = null,
    string? LastAuthor = null,
    int? PageCount = null,
    int? WordCount = null,
    int? CharacterCount = null,
    bool GrowAutofit = false,
    string? RevisionNumber = null);

internal sealed record DocPlainTextBookmark(string Name, int StartCp, int EndCp);

internal sealed record DocPlainTextPicture(int Cp, DocInlinePicture Picture);

internal sealed record DocPlainTextFloatingPicture(int Cp, DocInlinePicture Picture,
    int LeftTwips, int TopTwips, byte WrapCode = 2, bool BehindText = false,
    byte WrapSide = 0, byte HorizontalOrigin = 0, byte VerticalOrigin = 0,
    byte HorizontalAlignment = 0, byte VerticalAlignment = 0,
    int DistanceTopEmu = 0, int DistanceBottomEmu = 0,
    int DistanceLeftEmu = 114300, int DistanceRightEmu = 114300);

internal sealed record DocPlainTextFormatRun(int Start, int End,
    DocCharacterFormatting Formatting);

internal sealed record DocPlainTextParagraphStyleRun(int Start, int End, int StyleIndex,
    DocParagraphFormatting? Formatting = null);
