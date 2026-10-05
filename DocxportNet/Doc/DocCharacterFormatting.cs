using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Width and grouping ID for a manually fitted text region.</summary>
public sealed record DocFitText(int WidthTwips, int Id);

/// <summary>Direct character properties; null means the CHPX does not set that property.</summary>
public sealed record DocCharacterFormatting(bool? Bold, bool? Italic, ushort? SizeHalfPoints,
    int? CharacterStyleIndex = null, bool? Strike = null, bool? Caps = null,
    bool? SmallCaps = null, bool? Hidden = null, byte? UnderlineCode = null,
    uint? ColorRef = null, byte? HighlightCode = null, byte? ScriptCode = null,
    int? AsciiFontIndex = null, int? EastAsiaFontIndex = null,
    int? HighAnsiFontIndex = null, string? AsciiFontName = null,
    string? EastAsiaFontName = null, string? HighAnsiFontName = null,
    bool? Special = null, short? CharacterSpacingTwips = null,
    ushort? LanguageId = null, ushort? EastAsiaLanguageId = null,
    ushort? ComplexScriptLanguageId = null,
    int? PictureDataOffset = null, bool? FieldData = null,
    bool? BoldStyleToggle = null, bool? ItalicStyleToggle = null,
    bool? StrikeStyleToggle = null, bool? CapsStyleToggle = null,
    bool? SmallCapsStyleToggle = null, bool? HiddenStyleToggle = null,
    uint? UnderlineColorRef = null, DocCellShading? Shading = null,
    int? ComplexScriptFontIndex = null, string? ComplexScriptFontName = null,
    bool? ComplexScriptBold = null, bool? ComplexScriptItalic = null,
    ushort? ComplexScriptSizeHalfPoints = null,
    bool? ComplexScriptBoldStyleToggle = null,
    bool? ComplexScriptItalicStyleToggle = null,
    bool? RightToLeftText = null, bool? ForceComplexScript = null,
    bool? RightToLeftStyleToggle = null,
    bool? ForceComplexScriptStyleToggle = null,
    bool? DeletedRevision = null, bool? InsertedRevision = null,
    int? DeletedRevisionAuthorIndex = null, int? InsertedRevisionAuthorIndex = null,
    string? DeletedRevisionAuthor = null, string? InsertedRevisionAuthor = null,
    DateTime? DeletedRevisionAt = null, DateTime? InsertedRevisionAt = null,
    ushort? KerningThresholdHalfPoints = null, bool? DoubleStrike = null,
    bool? DoubleStrikeStyleToggle = null,
    bool? Shadow = null, bool? Outline = null, bool? Emboss = null,
    bool? Imprint = null, bool? ShadowStyleToggle = null,
    bool? OutlineStyleToggle = null, bool? EmbossStyleToggle = null,
    bool? ImprintStyleToggle = null, ushort? CharacterScalePercent = null,
    short? BaselineOffsetHalfPoints = null, DocParagraphBorder? Border = null,
    int? SymbolFontIndex = null, string? SymbolFontName = null,
    ushort? SymbolCharacter = null, bool? SnapToGrid = null,
    bool? SnapToGridStyleToggle = null, byte? EmphasisMarkCode = null,
    DocFitText? FitText = null)
{
    public static DocCharacterFormatting Empty { get; } = new(null, null, null);
    public bool IsEmpty => Bold == null && Italic == null && SizeHalfPoints == null &&
        CharacterStyleIndex == null && Strike == null && Caps == null &&
        SmallCaps == null && Hidden == null && UnderlineCode == null &&
        ColorRef == null && HighlightCode == null && ScriptCode == null &&
        AsciiFontIndex == null && EastAsiaFontIndex == null &&
        HighAnsiFontIndex == null && AsciiFontName == null &&
        EastAsiaFontName == null && HighAnsiFontName == null && Special == null &&
        CharacterSpacingTwips == null && LanguageId == null &&
        EastAsiaLanguageId == null && ComplexScriptLanguageId == null &&
        PictureDataOffset == null && FieldData == null && BoldStyleToggle == null &&
        ItalicStyleToggle == null && StrikeStyleToggle == null &&
        CapsStyleToggle == null && SmallCapsStyleToggle == null &&
        HiddenStyleToggle == null && UnderlineColorRef == null && Shading == null &&
        ComplexScriptFontIndex == null && ComplexScriptFontName == null &&
        ComplexScriptBold == null && ComplexScriptItalic == null &&
        ComplexScriptSizeHalfPoints == null &&
        ComplexScriptBoldStyleToggle == null &&
        ComplexScriptItalicStyleToggle == null &&
        RightToLeftText == null && ForceComplexScript == null &&
        RightToLeftStyleToggle == null && ForceComplexScriptStyleToggle == null &&
        DeletedRevision == null && InsertedRevision == null &&
        DeletedRevisionAuthorIndex == null && InsertedRevisionAuthorIndex == null &&
        DeletedRevisionAuthor == null && InsertedRevisionAuthor == null &&
        DeletedRevisionAt == null && InsertedRevisionAt == null &&
        KerningThresholdHalfPoints == null && DoubleStrike == null &&
        DoubleStrikeStyleToggle == null && Shadow == null && Outline == null &&
        Emboss == null && Imprint == null && ShadowStyleToggle == null &&
        OutlineStyleToggle == null && EmbossStyleToggle == null &&
        ImprintStyleToggle == null && CharacterScalePercent == null &&
        BaselineOffsetHalfPoints == null && Border == null &&
        SymbolFontIndex == null && SymbolFontName == null && SymbolCharacter == null &&
        SnapToGrid == null && SnapToGridStyleToggle == null &&
        EmphasisMarkCode == null && FitText == null;

    internal bool HasStyleToggles => BoldStyleToggle != null || ItalicStyleToggle != null ||
        StrikeStyleToggle != null || DoubleStrikeStyleToggle != null ||
        CapsStyleToggle != null ||
        SmallCapsStyleToggle != null || HiddenStyleToggle != null ||
        ComplexScriptBoldStyleToggle != null ||
        ComplexScriptItalicStyleToggle != null ||
        RightToLeftStyleToggle != null ||
        ForceComplexScriptStyleToggle != null || ShadowStyleToggle != null ||
        OutlineStyleToggle != null || EmbossStyleToggle != null ||
        ImprintStyleToggle != null || SnapToGridStyleToggle != null;

    internal DocCharacterFormatting ResolveStyleToggles(DocCharacterFormatting style) =>
        this with
        {
            Bold = BoldStyleToggle is bool bold ? (style.Bold ?? false) ^ bold : Bold,
            Italic = ItalicStyleToggle is bool italic ? (style.Italic ?? false) ^ italic : Italic,
            ComplexScriptBold = ComplexScriptBoldStyleToggle is bool csBold
                ? (style.ComplexScriptBold ?? false) ^ csBold : ComplexScriptBold,
            ComplexScriptItalic = ComplexScriptItalicStyleToggle is bool csItalic
                ? (style.ComplexScriptItalic ?? false) ^ csItalic : ComplexScriptItalic,
            RightToLeftText = RightToLeftStyleToggle is bool rtl
                ? (style.RightToLeftText ?? false) ^ rtl : RightToLeftText,
            ForceComplexScript = ForceComplexScriptStyleToggle is bool cs
                ? (style.ForceComplexScript ?? false) ^ cs : ForceComplexScript,
            Strike = StrikeStyleToggle is bool strike ? (style.Strike ?? false) ^ strike : Strike,
            DoubleStrike = DoubleStrikeStyleToggle is bool doubleStrike
                ? (style.DoubleStrike ?? false) ^ doubleStrike : DoubleStrike,
            Shadow = ShadowStyleToggle is bool shadow
                ? (style.Shadow ?? false) ^ shadow : Shadow,
            Outline = OutlineStyleToggle is bool outline
                ? (style.Outline ?? false) ^ outline : Outline,
            Emboss = EmbossStyleToggle is bool emboss
                ? (style.Emboss ?? false) ^ emboss : Emboss,
            Imprint = ImprintStyleToggle is bool imprint
                ? (style.Imprint ?? false) ^ imprint : Imprint,
            Caps = CapsStyleToggle is bool caps ? (style.Caps ?? false) ^ caps : Caps,
            SmallCaps = SmallCapsStyleToggle is bool smallCaps
                ? (style.SmallCaps ?? false) ^ smallCaps : SmallCaps,
            Hidden = HiddenStyleToggle is bool hidden ? (style.Hidden ?? false) ^ hidden : Hidden,
            SnapToGrid = SnapToGridStyleToggle is bool snap
                ? (style.SnapToGrid ?? true) ^ snap : SnapToGrid,
            BoldStyleToggle = null,
            ItalicStyleToggle = null,
            StrikeStyleToggle = null,
            DoubleStrikeStyleToggle = null,
            ShadowStyleToggle = null,
            OutlineStyleToggle = null,
            EmbossStyleToggle = null,
            ImprintStyleToggle = null,
            CapsStyleToggle = null,
            SmallCapsStyleToggle = null,
            HiddenStyleToggle = null,
            ComplexScriptBoldStyleToggle = null,
            ComplexScriptItalicStyleToggle = null,
            RightToLeftStyleToggle = null,
            ForceComplexScriptStyleToggle = null,
            SnapToGridStyleToggle = null
        };
}

public sealed record DocCharacterFormattingRange(uint CpStart, uint CpEnd,
    DocCharacterFormatting Formatting);

internal static class DocCharacterFormattingReader
{
    public static IReadOnlyList<DocCharacterFormattingRange> Read(DocTextIndex index)
    {
        var ranges = new List<DocCharacterFormattingRange>();
        var pieceProperties = new DocPiecePropertyReader(index);
        IReadOnlyList<DocParagraphStyleRange>? paragraphRanges = null;
        IReadOnlyDictionary<int, DocStyleDefinition>? styles = null;
        foreach (var page in index.FormattingPages.Where(x => x.IsCharacterFormatting))
        foreach (var run in page.Runs.Where(x => x.Kind == "ChpxRange"))
        {
            var property = run.Children.FirstOrDefault(x => x.Kind == "Chpx");
            var chpx = property?.Offset is long offset && property.Length is long length
                ? index.Structure.ReadRange("WordDocument", offset, checked((int)length))
                : Array.Empty<byte>();
            if (chpx.Length != 0 && chpx.Length != chpx[0] + 1)
                throw new InvalidDataException("A CHPX has an invalid length.");
            var directSprms = chpx.Length == 0 ? Array.Empty<byte>() :
                chpx.AsSpan(1).ToArray();
            var fcStart = long.Parse(run.Attributes["fcStart"], CultureInfo.InvariantCulture);
            var fcEnd = long.Parse(run.Attributes["fcEnd"], CultureInfo.InvariantCulture);
            foreach (var piece in index.Pieces)
            {
                var start = Math.Max(fcStart, piece.TextOffset);
                var end = Math.Min(fcEnd, piece.TextOffset + piece.TextByteLength);
                if (start >= end) continue;
                var pieceSprms = pieceProperties.Read(piece, character: true);
                var sprms = new byte[directSprms.Length + pieceSprms.Length];
                directSprms.CopyTo(sprms, 0);
                pieceSprms.CopyTo(sprms, directSprms.Length);
                var formatting = ParseSprms(sprms);
                if (formatting.DeletedRevisionAuthorIndex is int deletedAuthor)
                    formatting = formatting with { DeletedRevisionAuthor =
                        deletedAuthor >= 0 && deletedAuthor < index.RevisionAuthors.Count
                            ? index.RevisionAuthors[deletedAuthor] : "Unknown" };
                if (formatting.InsertedRevisionAuthorIndex is int insertedAuthor)
                    formatting = formatting with { InsertedRevisionAuthor =
                        insertedAuthor >= 0 && insertedAuthor < index.RevisionAuthors.Count
                            ? index.RevisionAuthors[insertedAuthor] : "Unknown" };
                if (formatting.AsciiFontIndex != null ||
                    formatting.EastAsiaFontIndex != null ||
                    formatting.HighAnsiFontIndex != null ||
                    formatting.ComplexScriptFontIndex != null ||
                    formatting.SymbolFontIndex != null)
                    formatting = ResolveFonts(formatting, index.Fonts);
                if (formatting.IsEmpty) continue;
                var width = piece.Encoding == "utf16" ? 2 : 1;
                if ((start - piece.TextOffset) % width != 0 ||
                    (end - piece.TextOffset) % width != 0)
                    throw new InvalidDataException("A CHPX boundary splits a character.");
                var cpStart = piece.GetCharacterPosition(start);
                var cpEnd = piece.GetCharacterPosition(end);
                if (!formatting.HasStyleToggles)
                {
                    ranges.Add(new DocCharacterFormattingRange(cpStart, cpEnd,
                        formatting));
                    continue;
                }
                paragraphRanges ??= index.ParagraphStyles;
                styles ??= index.StyleDefinitions.ToDictionary(x => x.Index);
                var cuts = paragraphRanges.Where(x => x.CpEnd > cpStart &&
                    x.CpStart < cpEnd).SelectMany(x => new[] { x.CpStart, x.CpEnd })
                    .Where(x => x > cpStart && x < cpEnd)
                    .Append(cpStart).Append(cpEnd).Distinct().OrderBy(x => x).ToArray();
                for (var i = 0; i < cuts.Length - 1; i++)
                {
                    var cp = cuts[i];
                    var paragraph = paragraphRanges.LastOrDefault(x =>
                        x.CpStart <= cp && cp < x.CpEnd);
                    var paragraphStyle = paragraph != null &&
                        styles.TryGetValue(paragraph.StyleIndex, out var foundParagraph)
                        ? foundParagraph.CharacterFormatting : DocCharacterFormatting.Empty;
                    var characterStyle = formatting.CharacterStyleIndex is int styleIndex &&
                        styles.TryGetValue(styleIndex, out var foundCharacter) &&
                        foundCharacter.Type == 2
                        ? foundCharacter.CharacterFormatting : DocCharacterFormatting.Empty;
                    var appliedStyle = OverlayToggleBase(paragraphStyle, characterStyle);
                    ranges.Add(new DocCharacterFormattingRange(cp, cuts[i + 1],
                        formatting.ResolveStyleToggles(appliedStyle)));
                }
            }
        }
        return ranges.OrderBy(x => x.CpStart).ToArray();
    }

    internal static DocCharacterFormatting OverlayToggleBase(
        DocCharacterFormatting paragraphStyle, DocCharacterFormatting characterStyle) =>
        paragraphStyle with
        {
            // Boolean effects in a character style toggle the paragraph-style
            // state before a relative CHPX operand is evaluated.
            Bold = characterStyle.Bold is bool bold
                ? (paragraphStyle.Bold ?? false) ^ bold : paragraphStyle.Bold,
            Italic = characterStyle.Italic is bool italic
                ? (paragraphStyle.Italic ?? false) ^ italic : paragraphStyle.Italic,
            Strike = characterStyle.Strike is bool strike
                ? (paragraphStyle.Strike ?? false) ^ strike : paragraphStyle.Strike,
            DoubleStrike = characterStyle.DoubleStrike is bool doubleStrike
                ? (paragraphStyle.DoubleStrike ?? false) ^ doubleStrike
                : paragraphStyle.DoubleStrike,
            Shadow = characterStyle.Shadow is bool shadow
                ? (paragraphStyle.Shadow ?? false) ^ shadow : paragraphStyle.Shadow,
            Outline = characterStyle.Outline is bool outline
                ? (paragraphStyle.Outline ?? false) ^ outline : paragraphStyle.Outline,
            Emboss = characterStyle.Emboss is bool emboss
                ? (paragraphStyle.Emboss ?? false) ^ emboss : paragraphStyle.Emboss,
            Imprint = characterStyle.Imprint is bool imprint
                ? (paragraphStyle.Imprint ?? false) ^ imprint : paragraphStyle.Imprint,
            Caps = characterStyle.Caps ?? paragraphStyle.Caps,
            SmallCaps = characterStyle.SmallCaps ?? paragraphStyle.SmallCaps,
            Hidden = characterStyle.Hidden ?? paragraphStyle.Hidden,
            SnapToGrid = characterStyle.SnapToGrid ?? paragraphStyle.SnapToGrid,
            ComplexScriptBold = characterStyle.ComplexScriptBold ??
                paragraphStyle.ComplexScriptBold,
            ComplexScriptItalic = characterStyle.ComplexScriptItalic ??
                paragraphStyle.ComplexScriptItalic,
            RightToLeftText = characterStyle.RightToLeftText ??
                paragraphStyle.RightToLeftText,
            ForceComplexScript = characterStyle.ForceComplexScript ??
                paragraphStyle.ForceComplexScript
        };

    internal static DocCharacterFormatting ParseSprms(ReadOnlySpan<byte> sprms)
    {
        bool? bold = null, italic = null, strike = null, doubleStrike = null,
            shadow = null, outline = null, emboss = null, imprint = null, caps = null,
            smallCaps = null, hidden = null, special = null, fieldData = null,
            snapToGrid = null, snapToGridStyleToggle = null;
        bool? boldStyleToggle = null, italicStyleToggle = null,
            strikeStyleToggle = null, doubleStrikeStyleToggle = null,
            shadowStyleToggle = null, outlineStyleToggle = null,
            embossStyleToggle = null, imprintStyleToggle = null, capsStyleToggle = null,
            smallCapsStyleToggle = null, hiddenStyleToggle = null,
            complexBoldStyleToggle = null, complexItalicStyleToggle = null;
        bool? complexBold = null, complexItalic = null,
            rightToLeft = null, forceComplexScript = null,
            rightToLeftToggle = null, forceComplexScriptToggle = null;
        bool? deletedRevision = null, insertedRevision = null;
        int? deletedRevisionAuthorIndex = null, insertedRevisionAuthorIndex = null;
        DateTime? deletedRevisionAt = null, insertedRevisionAt = null;
        ushort? complexSize = null;
        static (bool? Absolute, bool? Relative) DecodeToggle(byte value) => value switch
        {
            0 => (false, null),
            1 => (true, null),
            0x80 => (null, false),
            0x81 => (null, true),
            _ => throw new InvalidDataException("A character toggle operand is invalid.")
        };
        ushort? size = null, kerningThreshold = null, characterScale = null;
        short? baselineOffset = null;
        byte? underlineCode = null;
        uint? colorRef = null;
        var modernColorSeen = false;
        byte? highlightCode = null, scriptCode = null, emphasisMarkCode = null;
        uint? underlineColorRef = null;
        DocCellShading? shading = null;
        var modernShadingSeen = false;
        DocParagraphBorder? border = null;
        DocFitText? fitText = null;
        var modernBorderSeen = false;
        int? asciiFont = null, eastAsiaFont = null, highAnsiFont = null,
            complexScriptFont = null;
        int? styleIndex = null;
        short? characterSpacing = null;
        ushort? languageId = null, eastAsiaLanguageId = null, complexScriptLanguageId = null;
        int? pictureDataOffset = null, symbolFontIndex = null;
        ushort? symbolCharacter = null;
        for (var offset = 0; offset < sprms.Length;)
        {
            if (offset + 2 > sprms.Length) throw new InvalidDataException("A character SPRM is truncated.");
            var sprm = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
            offset += 2;
            var spra = sprm >> 13;
            var operandLength = spra switch
            {
                0 or 1 => 1,
                2 or 4 or 5 => 2,
                3 => 4,
                6 when offset < sprms.Length => 1 + sprms[offset],
                7 => 3,
                _ => throw new InvalidDataException("A character SPRM has an invalid operand.")
            };
            if (offset + operandLength > sprms.Length)
                throw new InvalidDataException("A character SPRM operand is truncated.");
            switch (sprm)
            {
                case 0x0800:
                    var (deletedAbsolute, deletedToggle) = DecodeToggle(sprms[offset]);
                    deletedRevision = deletedAbsolute ?? (deletedRevision ?? false) ^ deletedToggle!.Value;
                    break;
                case 0x0801:
                    var (insertedAbsolute, insertedToggle) = DecodeToggle(sprms[offset]);
                    insertedRevision = insertedAbsolute ?? (insertedRevision ?? false) ^ insertedToggle!.Value;
                    break;
                case 0x4804:
                    insertedRevisionAuthorIndex = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    break;
                case 0x6805:
                    insertedRevisionAt = DocRevisionDate.Decode(BinaryPrimitives.ReadUInt32LittleEndian(
                        sprms.Slice(offset)));
                    break;
                case 0x4863:
                    deletedRevisionAuthorIndex = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    break;
                case 0x6864:
                    deletedRevisionAt = DocRevisionDate.Decode(BinaryPrimitives.ReadUInt32LittleEndian(
                        sprms.Slice(offset)));
                    break;
                case 0xCA76:
                    if (operandLength != 9 || sprms[offset] != 8)
                        throw new InvalidDataException("A DOC fit-text operand is invalid.");
                    fitText = new DocFitText(
                        BinaryPrimitives.ReadInt32LittleEndian(sprms.Slice(offset + 1)),
                        BinaryPrimitives.ReadInt32LittleEndian(sprms.Slice(offset + 5)));
                    break;
                case 0x2A34:
                    if (sprms[offset] > 4)
                        throw new InvalidDataException("A DOC emphasis mark is invalid.");
                    emphasisMarkCode = sprms[offset];
                    break;
                case 0x2A33:
                    if (sprms[offset] != 0)
                        throw new InvalidDataException("A character plain-reset operand is invalid.");
                    bold = italic = strike = doubleStrike = shadow = outline =
                        emboss = imprint = caps = smallCaps = hidden = null;
                    complexBold = complexItalic = null;
                    boldStyleToggle = italicStyleToggle = strikeStyleToggle =
                        doubleStrikeStyleToggle = null;
                    shadowStyleToggle = outlineStyleToggle = embossStyleToggle =
                        imprintStyleToggle = null;
                    capsStyleToggle = smallCapsStyleToggle = hiddenStyleToggle = null;
                    complexBoldStyleToggle = complexItalicStyleToggle = null;
                    size = complexSize = kerningThreshold = characterScale = null;
                    emphasisMarkCode = null;
                    fitText = null;
                    baselineOffset = null;
                    underlineCode = scriptCode = null;
                    colorRef = underlineColorRef = null;
                    modernColorSeen = false;
                    shading = null;
                    modernShadingSeen = false;
                    border = null;
                    modernBorderSeen = false;
                    asciiFont = eastAsiaFont = highAnsiFont = complexScriptFont = null;
                    characterSpacing = null;
                    languageId = eastAsiaLanguageId = complexScriptLanguageId = null;
                    styleIndex = null;
                    break;
                case 0x0835: (bold, boldStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0836: (italic, italicStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x085C: (complexBold, complexBoldStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x085D: (complexItalic, complexItalicStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x085A: (rightToLeft, rightToLeftToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0882: (forceComplexScript, forceComplexScriptToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0837: (strike, strikeStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x2A53: (doubleStrike, doubleStrikeStyleToggle) =
                    DecodeToggle(sprms[offset]); break;
                case 0x0838: (outline, outlineStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0839: (shadow, shadowStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0858: (emboss, embossStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0854: (imprint, imprintStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x083A: (smallCaps, smallCapsStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x083B: (caps, capsStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x083C: (hidden, hiddenStyleToggle) = DecodeToggle(sprms[offset]); break;
                case 0x0855: special = sprms[offset] switch { 0 => false, 1 => true, _ => null }; break;
                case 0x0806: fieldData = sprms[offset] != 0; break;
                case 0x0868: (snapToGrid, snapToGridStyleToggle) =
                    DecodeToggle(sprms[offset]); break;
                case 0x6A09:
                    symbolFontIndex = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
                    symbolCharacter = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset + 2));
                    break;
                case 0x6A03: pictureDataOffset = BinaryPrimitives.ReadInt32LittleEndian(
                    sprms.Slice(offset)); break;
                case 0x2A3E: underlineCode = sprms[offset]; break;
                case 0x2A42:
                    if (!modernColorSeen) colorRef = IndexedColor(sprms[offset]);
                    break;
                case 0x6870:
                    colorRef = BinaryPrimitives.ReadUInt32LittleEndian(sprms.Slice(offset));
                    modernColorSeen = true;
                    break;
                case 0x6877: underlineColorRef = BinaryPrimitives.ReadUInt32LittleEndian(
                    sprms.Slice(offset)); break;
                case 0x2A0C: highlightCode = sprms[offset]; break;
                case 0x4866 when !modernShadingSeen:
                    var shd80 = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
                    var foregroundIndex = shd80 & 31;
                    var backgroundIndex = (shd80 >> 5) & 31;
                    var legacyPattern = (ushort)(shd80 >> 10);
                    if (foregroundIndex == 31 && backgroundIndex == 31 &&
                        legacyPattern == 63)
                    {
                        shading = new DocCellShading(null, null, 0);
                        break;
                    }
                    if (foregroundIndex > 16 || backgroundIndex > 16 ||
                        DocShadingPatterns.ToOpenXml(legacyPattern) == null)
                        throw new InvalidDataException("A legacy character shading value is invalid.");
                    shading = new DocCellShading(backgroundIndex == 0 ? null :
                        IndexedColor((byte)backgroundIndex),
                        foregroundIndex == 0 ? null : IndexedColor((byte)foregroundIndex),
                        legacyPattern);
                    break;
                case 0xCA71 when operandLength == 11 && sprms[offset] == 10:
                    var foreground = BinaryPrimitives.ReadUInt32LittleEndian(
                        sprms.Slice(offset + 1));
                    var background = BinaryPrimitives.ReadUInt32LittleEndian(
                        sprms.Slice(offset + 5));
                    var pattern = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 9));
                    if (foreground == 0xFFFFFFFFu &&
                        background == 0xFF000000u && pattern == 0)
                        shading = new DocCellShading(null, null, ushort.MaxValue);
                    else if (DocShadingPatterns.ToOpenXml(pattern) != null)
                        shading = new DocCellShading((background >> 24) == 0
                            ? background & 0xFFFFFF : null,
                            (foreground >> 24) == 0
                            ? foreground & 0xFFFFFF : null, pattern);
                    modernShadingSeen = true;
                    break;
                case 0x6865 when !modernBorderSeen:
                    border = DocBinaryCompat.AllBytesEqual(sprms.Slice(offset, 4), 0xFF)
                        ? new DocParagraphBorder(0, 0, 0, null)
                        : DocParagraphBorder.Parse80(sprms.Slice(offset, 4));
                    break;
                case 0xCA72 when operandLength == 9 && sprms[offset] == 8:
                    border = DocBinaryCompat.AllBytesEqual(sprms.Slice(offset + 1, 8), 0xFF)
                        ? new DocParagraphBorder(0, 0, 0, null)
                        : DocParagraphBorder.Parse(sprms.Slice(offset, 9));
                    modernBorderSeen = true;
                    break;
                case 0x2A48: scriptCode = sprms[offset]; break;
                case 0x4A4F: asciiFont = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4A50: eastAsiaFont = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4A51: highAnsiFont = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4A5E: complexScriptFont = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4A43: size = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x484B: kerningThreshold = BinaryPrimitives.ReadUInt16LittleEndian(
                    sprms.Slice(offset)); break;
                case 0x4852:
                    var scale = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
                    if (scale is < 1 or > 600)
                        throw new InvalidDataException("A character scale is outside 1–600 percent.");
                    characterScale = scale;
                    break;
                case 0x4845:
                    var position = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    if (position is < -3168 or > 3168)
                        throw new InvalidDataException("A baseline offset is outside ±3168 half-points.");
                    baselineOffset = position;
                    break;
                case 0x4A61: complexSize = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4A30: styleIndex = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x8840: characterSpacing = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x486D or 0x4873:
                    languageId = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x486E or 0x4874:
                    eastAsiaLanguageId = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x485F:
                    complexScriptLanguageId = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
            }
            offset += operandLength;
        }
        return new DocCharacterFormatting(bold, italic, size, styleIndex,
            strike, caps, smallCaps, hidden, underlineCode, colorRef,
            highlightCode, scriptCode, asciiFont, eastAsiaFont, highAnsiFont,
            Special: special, CharacterSpacingTwips: characterSpacing,
            LanguageId: languageId, EastAsiaLanguageId: eastAsiaLanguageId,
            ComplexScriptLanguageId: complexScriptLanguageId,
            PictureDataOffset: pictureDataOffset, FieldData: fieldData,
            BoldStyleToggle: boldStyleToggle,
            ItalicStyleToggle: italicStyleToggle,
            StrikeStyleToggle: strikeStyleToggle,
            DoubleStrike: doubleStrike,
            DoubleStrikeStyleToggle: doubleStrikeStyleToggle,
            Shadow: shadow, Outline: outline, Emboss: emboss, Imprint: imprint,
            ShadowStyleToggle: shadowStyleToggle,
            OutlineStyleToggle: outlineStyleToggle,
            EmbossStyleToggle: embossStyleToggle,
            ImprintStyleToggle: imprintStyleToggle,
            CharacterScalePercent: characterScale,
            BaselineOffsetHalfPoints: baselineOffset,
            CapsStyleToggle: capsStyleToggle,
            SmallCapsStyleToggle: smallCapsStyleToggle,
            HiddenStyleToggle: hiddenStyleToggle,
            UnderlineColorRef: underlineColorRef, Shading: shading,
            ComplexScriptFontIndex: complexScriptFont,
            ComplexScriptBold: complexBold, ComplexScriptItalic: complexItalic,
            ComplexScriptSizeHalfPoints: complexSize,
            ComplexScriptBoldStyleToggle: complexBoldStyleToggle,
            ComplexScriptItalicStyleToggle: complexItalicStyleToggle,
            RightToLeftText: rightToLeft,
            ForceComplexScript: forceComplexScript,
            RightToLeftStyleToggle: rightToLeftToggle,
            ForceComplexScriptStyleToggle: forceComplexScriptToggle,
            DeletedRevision: deletedRevision, InsertedRevision: insertedRevision,
            DeletedRevisionAuthorIndex: deletedRevisionAuthorIndex,
            InsertedRevisionAuthorIndex: insertedRevisionAuthorIndex,
            DeletedRevisionAt: deletedRevisionAt, InsertedRevisionAt: insertedRevisionAt,
            KerningThresholdHalfPoints: kerningThreshold, Border: border,
            SymbolFontIndex: symbolFontIndex, SymbolCharacter: symbolCharacter,
            SnapToGrid: snapToGrid, SnapToGridStyleToggle: snapToGridStyleToggle,
            EmphasisMarkCode: emphasisMarkCode, FitText: fitText);
    }

    private static uint IndexedColor(byte index) => index switch
    {
        0 => 0xFF000000u, // automatic
        1 => 0x000000u,
        2 => 0xFF0000u, // blue
        3 => 0xFFFF00u,
        4 => 0x00FF00u,
        5 => 0xFF00FFu,
        6 => 0x0000FFu, // red
        7 => 0x00FFFFu,
        8 => 0xFFFFFFu,
        9 => 0x800000u,
        10 => 0x808000u,
        11 => 0x008000u,
        12 or 13 => 0x800080u,
        14 => 0x008080u,
        15 => 0x808080u,
        16 => 0xC0C0C0u,
        _ => throw new InvalidDataException("A DOC indexed text color is invalid.")
    };

    internal static DocCharacterFormatting ResolveFonts(DocCharacterFormatting formatting,
        IReadOnlyList<DocFontDefinition> fonts)
    {
        string? Name(int? index) => index is int i && i >= 0 && i < fonts.Count
            ? fonts[i].Name : null;
        return formatting with
        {
            AsciiFontName = Name(formatting.AsciiFontIndex),
            EastAsiaFontName = Name(formatting.EastAsiaFontIndex),
            HighAnsiFontName = Name(formatting.HighAnsiFontIndex),
            ComplexScriptFontName = Name(formatting.ComplexScriptFontIndex),
            SymbolFontName = Name(formatting.SymbolFontIndex)
        };
    }
}
