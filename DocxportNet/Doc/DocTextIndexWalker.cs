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
    private int BytesPerCharacter => Encoding switch
    {
        "compressed" => 1,
        "utf16" => 2,
        _ => throw new InvalidDataException("The text piece has an unknown encoding.")
    };

    /// <summary>Maps a position in this piece to its physical WordDocument offset.</summary>
    public long GetTextOffset(uint cp)
    {
        if (cp < CpStart || cp > CpEnd)
            throw new ArgumentOutOfRangeException(nameof(cp));
        return checked(TextOffset + (long)(cp - CpStart) * BytesPerCharacter);
    }

    /// <summary>Maps a character-aligned physical offset back to this piece.</summary>
    public uint GetCharacterPosition(long textOffset)
    {
        var difference = textOffset - TextOffset;
        if (difference < 0 || difference > TextByteLength ||
            difference % BytesPerCharacter != 0)
            throw new ArgumentOutOfRangeException(nameof(textOffset));
        var cp = checked(CpStart + (uint)(difference / BytesPerCharacter));
        if (cp > CpEnd) throw new InvalidDataException(
            "The text piece byte length exceeds its character range.");
        return cp;
    }

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
    public long TextOffsetStart => Piece.GetTextOffset(CpStart);
    public long TextOffsetEnd => Piece.GetTextOffset(CpEnd);
    public string Text => Piece.Text.Substring(checked((int)(CpStart - Piece.CpStart)),
        checked((int)(CpEnd - CpStart)));
}

/// <summary>Owns the source DOC while indexed text ranges may be materialized on demand.</summary>
public sealed class DocTextIndex : IDisposable
{
    private readonly Lazy<IReadOnlyList<DocIndexedFormattingPage>> _formattingPages;
    private readonly Lazy<IReadOnlyList<DocStructureNode>> _styles;
    private readonly Lazy<IReadOnlyList<DocStructureNode>> _sections;
    private readonly Lazy<IReadOnlyList<DocStorySection>> _storySections;
    private readonly Lazy<IReadOnlyList<DocHeaderStory>> _headerStories;
    private readonly Lazy<IReadOnlyList<DocCharacterFormattingRange>> _characterFormatting;
    private readonly Lazy<IReadOnlyList<DocStyleDefinition>> _styleDefinitions;
    private readonly Lazy<IReadOnlyList<DocParagraphStyleRange>> _paragraphStyles;
    private readonly Lazy<IReadOnlyList<DocFontDefinition>> _fonts;
    private readonly Lazy<DocCharacterFormatting> _defaultCharacterFormatting;
    private readonly Lazy<IReadOnlyList<DocBookmark>> _bookmarks;
    private readonly Lazy<IReadOnlyList<DocFloatingPicture>> _floatingPictures;
    private readonly Lazy<DocListIndex> _lists;
    private readonly Lazy<IReadOnlyList<string>> _revisionAuthors;

    internal DocTextIndex(DocStructure structure, IReadOnlyList<DocIndexedTextPiece> pieces)
    {
        Structure = structure;
        Pieces = pieces;
        _formattingPages = new Lazy<IReadOnlyList<DocIndexedFormattingPage>>(LoadFormattingPages);
        _styles = new Lazy<IReadOnlyList<DocStructureNode>>(LoadStyles);
        _sections = new Lazy<IReadOnlyList<DocStructureNode>>(LoadSections);
        _storySections = new Lazy<IReadOnlyList<DocStorySection>>(LoadStorySections);
        _headerStories = new Lazy<IReadOnlyList<DocHeaderStory>>(LoadHeaderStories);
        _characterFormatting = new Lazy<IReadOnlyList<DocCharacterFormattingRange>>(
            () => DocCharacterFormattingReader.Read(this));
        _styleDefinitions = new Lazy<IReadOnlyList<DocStyleDefinition>>(
            () => DocStyleDefinitionsReader.Read(this));
        _paragraphStyles = new Lazy<IReadOnlyList<DocParagraphStyleRange>>(
            () => DocParagraphStyleReader.Read(this));
        _fonts = new Lazy<IReadOnlyList<DocFontDefinition>>(
            () => DocFontTable.Read(Structure));
        _defaultCharacterFormatting = new Lazy<DocCharacterFormatting>(LoadDefaultFonts);
        _bookmarks = new Lazy<IReadOnlyList<DocBookmark>>(() => DocBookmarkReader.Read(this));
        _floatingPictures = new Lazy<IReadOnlyList<DocFloatingPicture>>(() =>
            DocFloatingPictureReader.ReadMain(Structure)
                .Concat(DocFloatingPictureReader.ReadHeaders(Structure)).ToArray());
        _lists = new Lazy<DocListIndex>(() => DocListReader.Read(this));
        _revisionAuthors = new Lazy<IReadOnlyList<string>>(LoadRevisionAuthors);
    }

    public DocStructure Structure { get; }
    public IReadOnlyList<DocIndexedTextPiece> Pieces { get; }
    public IReadOnlyList<DocIndexedFormattingPage> FormattingPages => _formattingPages.Value;
    public IReadOnlyList<DocStructureNode> Styles => _styles.Value;
    public IReadOnlyList<DocStructureNode> Sections => _sections.Value;
    public IReadOnlyList<DocStorySection> StorySections => _storySections.Value;
    public IReadOnlyList<DocHeaderStory> HeaderStories => _headerStories.Value;
    public IReadOnlyList<DocCharacterFormattingRange> CharacterFormatting => _characterFormatting.Value;
    public IReadOnlyList<DocStyleDefinition> StyleDefinitions => _styleDefinitions.Value;
    public IReadOnlyList<DocParagraphStyleRange> ParagraphStyles => _paragraphStyles.Value;
    public IReadOnlyList<DocFontDefinition> Fonts => _fonts.Value;
    public DocCharacterFormatting DefaultCharacterFormatting => _defaultCharacterFormatting.Value;
    public IReadOnlyList<DocBookmark> Bookmarks => _bookmarks.Value;
    public IReadOnlyList<DocFloatingPicture> FloatingPictures => _floatingPictures.Value;
    public DocListIndex Lists => _lists.Value;
    public IReadOnlyList<string> RevisionAuthors => _revisionAuthors.Value;
    public bool? EvenAndOddHeaders
    {
        get
        {
            var location = Structure.Locations.FirstOrDefault(x => x.Name == "DocumentProperties" &&
                x.IsPresent);
            return location == null ? null :
                (Structure.ReadRange(location.StreamName, location.Offset, 1)[0] & 1) != 0;
        }
    }
    public short? DefaultTabStopTwips
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 12) return null;
            var bytes = Structure.ReadRange(location.StreamName, location.Offset + 10, 2);
            var value = System.Buffers.Binary.BinaryPrimitives.ReadInt16LittleEndian(bytes);
            return value <= 0 ? null : value;
        }
    }
    public bool? BalanceSingleByteDoubleByteWidth
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 10) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 9, 1)[0] & 0x80) == 0;
        }
    }
    public bool? ApplyBreakingRules
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 514) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 513, 1)[0]
                & 0x20) != 0;
        }
    }
    public bool? GrowAutofit
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 515) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 514, 1)[0]
                & 0x02) != 0;
        }
    }
    public bool? MirrorMargins
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 7) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 6, 1)[0] & 0x20) != 0;
        }
    }
    public bool? AutoHyphenation
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 6) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 5, 1)[0] & 0x10) != 0;
        }
    }
    public bool? HyphenateCaps
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 6) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 5, 1)[0] & 0x08) != 0;
        }
    }
    public short? HyphenationZoneTwips
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 16) return null;
            var bytes = Structure.ReadRange(location.StreamName, location.Offset + 14, 2);
            return System.Buffers.Binary.BinaryPrimitives.ReadInt16LittleEndian(bytes);
        }
    }
    public short? ConsecutiveHyphenLimit
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 18) return null;
            var bytes = Structure.ReadRange(location.StreamName, location.Offset + 16, 2);
            return System.Buffers.Binary.BinaryPrimitives.ReadInt16LittleEndian(bytes);
        }
    }
    public bool? GutterAtTop
    {
        get
        {
            var location = Structure.FindLocation("DocumentProperties");
            if (location?.IsPresent != true || location.Length < 84) return null;
            return (Structure.ReadRange(location.StreamName, location.Offset + 83, 1)[0] & 0x80) != 0;
        }
    }
    public bool IsFormattingIndexLoaded => _formattingPages.IsValueCreated;
    public bool IsStyleIndexLoaded => _styles.IsValueCreated;
    public bool IsSectionIndexLoaded => _sections.IsValueCreated;
    public IReadOnlyList<DocPartRange> Parts => Structure.Parts;

    public IReadOnlyCollection<uint> GetFieldMarks(string partName)
        => new HashSet<uint>(GetFieldRecords(partName).Select(x => x.Cp));

    /// <summary>Reads field PLC markers and parsed types in global CP space.</summary>
    public IReadOnlyList<DocIndexedFieldMark> GetFieldRecords(string partName)
    {
        var name = partName == "Main" ? "MainFields" :
            partName == "Headers" ? "HeaderFields" : null;
        var location = name == null ? null : FindLocationNode(name);
        if (location == null) return [];
        DocFieldNavigator.Expand(Structure, location);
        return location.Children.SelectMany(x => x.Children).Select(x =>
            new DocIndexedFieldMark(
                uint.Parse(x.Attributes["globalCp"], CultureInfo.InvariantCulture),
                Convert.ToByte(x.Attributes["character"].Substring(2), 16),
                x.Attributes.TryGetValue("fieldType", out var type)
                    ? byte.Parse(type, CultureInfo.InvariantCulture) : (byte)0)).ToArray();
    }

    /// <summary>Maps a character position to its physical text byte offset in WordDocument.</summary>
    public long GetTextOffset(uint cp)
    {
        var piece = Pieces.FirstOrDefault(x => x.CpStart <= cp && cp < x.CpEnd)
            ?? throw new ArgumentOutOfRangeException(nameof(cp),
                "The character position has no text piece.");
        return piece.GetTextOffset(cp);
    }

    // A CP at a piece boundary has two physical addresses when pieces are
    // stored apart. A range start belongs to the next piece; an end belongs
    // to the preceding piece. Keep the endpoint fallback for empty stories.
    private long? BoundaryOffset(uint cp, bool end)
    {
        var piece = Pieces.FirstOrDefault(x => end
            ? x.CpStart < cp && cp <= x.CpEnd
            : x.CpStart <= cp && cp < x.CpEnd)
            ?? Pieces.FirstOrDefault(x => x.CpStart == cp || x.CpEnd == cp);
        return piece?.GetTextOffset(cp);
    }
    /// <summary>Reads a story with only the formatting ranges touching its CP range.</summary>
    public DocIndexedStoryText ReadStory(string partName)
        => BindStory(DocStoryTextReader.Read(this, partName));

    internal DocIndexedStoryText BindStory(DocStoryText text)
    {
        var formatting = CharacterFormatting.Where(x => x.CpStart < text.CpEnd &&
            x.CpEnd > text.CpStart).ToArray();
        var paragraphStyles = ParagraphStyles.Where(x => x.CpStart < text.CpEnd &&
            x.CpEnd > text.CpStart).ToArray();
        var tableRows = paragraphStyles.Where(x =>
                DocStoryTableRow.IsRow(x.Formatting) &&
                x.CpStart >= text.CpStart && x.CpEnd <= text.CpEnd)
            .Select(x => new DocStoryTableRow(
                checked((int)(x.CpStart - text.CpStart)),
                checked((int)(x.CpEnd - text.CpStart)), x.StyleIndex,
                x.Formatting!, x.CpStart, x.CpEnd, GetTextOffset(x.CpStart)))
            .ToArray();
        long? StartOffset(uint cp) => BoundaryOffset(cp, end: false);
        long? EndOffset(uint cp) => BoundaryOffset(cp, end: true);
        var inlinePictures = text.Paragraphs.SelectMany(x => x.Atoms)
            .Where(x => x.Kind == DocStoryAtomKind.InlinePicture)
            .Select(x =>
            {
                var source = formatting.LastOrDefault(y => y.CpStart <= x.CpStart &&
                    x.CpStart < y.CpEnd);
                return source?.Formatting.PictureDataOffset is int offset
                    ? new DocStoryInlinePicture(
                        checked((int)(x.CpStart - text.CpStart)), offset,
                        x.CpStart, GetTextOffset(x.CpStart)) : null;
            })
            .Where(x => x != null).Select(x => x!)
            .ToDictionary(x => x.Cp);
        return new DocIndexedStoryText(text, formatting, paragraphStyles)
        {
            ParagraphRanges = text.Paragraphs.Select(x => new DocStoryParagraphRange(
                checked((int)(x.CpStart - text.CpStart)),
                checked((int)(x.CpEnd - text.CpStart)), x.End,
                x.CpStart, x.CpEnd, StartOffset(x.CpStart),
                EndOffset(x.CpEnd))).ToArray(),
            Runs = formatting.Select(x =>
            {
                var start = Math.Max(x.CpStart, text.CpStart);
                var end = Math.Min(x.CpEnd, text.CpEnd);
                return new DocStoryCharacterRun(
                    checked((int)(start - text.CpStart)),
                    checked((int)(end - text.CpStart)), x.Formatting,
                    start, end, StartOffset(start), EndOffset(end));
            }).ToArray(),
            StoryParagraphStyles = paragraphStyles
                .Where(x => !DocStoryTableRow.IsRow(x.Formatting)).Select(x =>
            {
                var start = Math.Max(x.CpStart, text.CpStart);
                var end = Math.Min(x.CpEnd, text.CpEnd);
                return new DocStoryParagraphStyleRun(
                    checked((int)(start - text.CpStart)),
                    checked((int)(end - text.CpStart)), x.StyleIndex,
                    x.Formatting, start, end, StartOffset(start),
                    EndOffset(end));
            }).ToArray(),
            TableRows = tableRows,
            ListParagraphs = paragraphStyles.Where(x =>
                x.Formatting?.ListOverrideIndex > 0).Select(x =>
            {
                var start = Math.Max(x.CpStart, text.CpStart);
                var end = Math.Min(x.CpEnd, text.CpEnd);
                return new DocStoryListParagraph(
                    checked((int)(start - text.CpStart)),
                    checked((int)(end - text.CpStart)),
                    x.Formatting!.ListOverrideIndex!.Value,
                    x.Formatting.ListLevel ?? 0, start, end,
                    StartOffset(start), EndOffset(end));
            }).ToArray(),
            InlinePictures = inlinePictures,
            FieldMarks = GetFieldRecords(text.Name).Where(x => x.Cp >= text.CpStart &&
                x.Cp < text.CpEnd).Select(x => new DocStoryFieldMark(
                    checked((int)(x.Cp - text.CpStart)), x.Kind, x.FieldType,
                    x.Cp, GetTextOffset(x.Cp)))
                .ToArray(),
            Bookmarks = Bookmarks.Select((bookmark, id) => (Bookmark: bookmark, Id: id))
                .Where(x => x.Bookmark.CpStart >= text.CpStart &&
                    x.Bookmark.CpEnd <= text.CpEnd)
                .Select(x => new DocStoryBookmark(x.Bookmark.Name,
                    checked((int)(x.Bookmark.CpStart - text.CpStart)),
                    checked((int)(x.Bookmark.CpEnd - text.CpStart)), x.Id,
                    x.Bookmark.CpStart, x.Bookmark.CpEnd,
                    StartOffset(x.Bookmark.CpStart),
                    EndOffset(x.Bookmark.CpEnd))).ToArray(),
            FloatingPictures = FloatingPictures.Where(x => x.Cp >= text.CpStart &&
                x.Cp < text.CpEnd)
                .Select(x => new DocStoryFloatingPicture(
                    checked((int)(x.Cp - text.CpStart)), x.Bytes, x.ContentType,
                    checked((x.RightTwips - x.LeftTwips) * 635L),
                    checked((x.BottomTwips - x.TopTwips) * 635L),
                    x.LeftTwips, x.TopTwips, x.WrapCode, x.BehindText,
                    x.WrapSide, x.HorizontalOrigin, x.VerticalOrigin,
                    x.HorizontalAlignment, x.VerticalAlignment,
                    x.DistanceTopEmu, x.DistanceBottomEmu,
                    x.DistanceLeftEmu, x.DistanceRightEmu)
                {
                    Crop = x.Crop, FlipHorizontal = x.FlipHorizontal,
                    FlipVertical = x.FlipVertical,
                    RotationDegrees = x.RotationDegrees, ShapeId = x.ShapeId,
                    SourceCp = x.Cp, SourceTextOffset = GetTextOffset(x.Cp)
                })
                .GroupBy(x => x.Cp).ToDictionary(x => x.Key, x => x.ToArray())
        };
    }

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

    private DocCharacterFormatting LoadDefaultFonts()
    {
        var location = Structure.FindLocation("StyleSheet");
        if (location?.IsPresent != true || location.Length < 20)
            return DocCharacterFormatting.Empty;
        var headerLength = Structure.ReadRange(location.StreamName,
            location.Offset, 2);
        var headerBytes = 2 + System.Buffers.Binary.BinaryPrimitives
            .ReadUInt16LittleEndian(headerLength);
        if (headerBytes > location.Length) return DocCharacterFormatting.Empty;
        var header = DocStyleSheetHeader.Read(Structure.ReadRange(location.StreamName,
            location.Offset, headerBytes));
        IReadOnlyList<DocFontDefinition> fonts;
        try { fonts = Fonts; }
        catch (InvalidDataException) { return DocCharacterFormatting.Empty; }
        string? Name(short index) => index >= 0 && index < fonts.Count
            ? fonts[index].Name : null;
        return DocCharacterFormatting.Empty with
        {
            AsciiFontName = Name(header.AsciiFontIndex),
            EastAsiaFontName = Name(header.EastAsiaFontIndex),
            HighAnsiFontName = Name(header.HighAnsiFontIndex),
            ComplexScriptFontName = Name(header.ComplexScriptFontIndex)
        };
    }

    private IReadOnlyList<string> LoadRevisionAuthors()
    {
        var location = FindLocationNode("RevisionAuthors");
        if (location == null) return Array.Empty<string>();
        DocStringTableNavigator.Expand(Structure, location);
        return location.Children.SelectMany(x => x.Children)
            .SelectMany(x => x.Children)
            .Where(x => x.Kind == "STTBEntry")
            .Select(x => x.Attributes.TryGetValue("text", out var name)
                ? name : string.Empty).ToArray();
    }

    private IReadOnlyList<DocStructureNode> LoadSections()
    {
        var location = FindLocationNode("Sections");
        if (location == null) return Array.Empty<DocStructureNode>();
        DocSectionNavigator.Expand(Structure, location);
        return location.Children.SelectMany(plc => plc.Children).ToArray();
    }

    private IReadOnlyList<DocStorySection> LoadStorySections()
    {
        var sections = Sections;
        var headers = HeaderStories;
        if (headers.Count != 0 && headers.Count != 6 + sections.Count * 6)
            throw new InvalidDataException("The DOC header story slots do not match its sections.");
        long? StartOffset(uint cp) => BoundaryOffset(cp, end: false);
        long? EndOffset(uint cp) => BoundaryOffset(cp, end: true);
        return sections.Select((node, index) =>
        {
            var start = uint.Parse(node.Attributes["cpStart"], CultureInfo.InvariantCulture);
            var end = uint.Parse(node.Attributes["cpEnd"], CultureInfo.InvariantCulture);
            var slots = Enumerable.Range(0, 6).Select(slot =>
            {
                if (headers.Count == 0) return new DocStorySectionSlot(slot, false);
                var header = headers[6 + index * 6 + slot];
                return new DocStorySectionSlot(slot, !header.IsEmpty,
                    header.CpStart, header.IsEmpty ? header.CpEnd : header.CpEnd - 1);
            }).ToArray();
            var sepxOffset = long.Parse(node.Attributes["sepxOffset"],
                CultureInfo.InvariantCulture);
            var result = new DocStorySection(checked((int)start), checked((int)end),
                DocSectionFormatting.Read(this, node), slots,
                StartOffset(start), EndOffset(end),
                sepxOffset < 0 ? null : sepxOffset);
            result.Validate();
            return result;
        }).ToArray();
    }

    private IReadOnlyList<DocHeaderStory> LoadHeaderStories()
    {
        var location = FindLocationNode("HeadersAndFooters");
        if (location == null) return Array.Empty<DocHeaderStory>();
        DocStoryPlcNavigator.Expand(Structure, location);
        return location.Children.SelectMany(plc => plc.Children)
            .Select((node, slot) => new DocHeaderStory(this, node, slot)).ToArray();
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
