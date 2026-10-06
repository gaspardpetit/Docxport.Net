using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>A custom DOC tab stop and its TBD alignment and leader values.</summary>
public sealed record DocTabStop(short PositionTwips, byte Alignment, byte Leader);
public sealed record DocTabClearRange(short PositionTwips, ushort ToleranceTwips);
public sealed record DocCellShading(uint? FillRgb, uint? ForegroundRgb, ushort Pattern);
public sealed record DocCellBorders(DocParagraphBorder? Top = null,
    DocParagraphBorder? Left = null, DocParagraphBorder? Bottom = null,
    DocParagraphBorder? Right = null, DocParagraphBorder? TopLeftToBottomRight = null,
    DocParagraphBorder? TopRightToBottomLeft = null);
public sealed record DocTableBorders(DocParagraphBorder? Top, DocParagraphBorder? Left,
    DocParagraphBorder? Bottom, DocParagraphBorder? Right,
    DocParagraphBorder? InsideHorizontal, DocParagraphBorder? InsideVertical);
public sealed record DocTablePreferredWidth(byte Unit, ushort Value);
public sealed record DocCellMargins(ushort? Top = null, ushort? Left = null,
    ushort? Bottom = null, ushort? Right = null);

/// <summary>Selected paragraph SPRMs; null means the property is not set here.</summary>
public sealed record DocParagraphFormatting(byte? Justification, ushort? BeforeTwips,
    ushort? AfterTwips, short? LeftTwips = null, short? RightTwips = null,
    short? FirstLineTwips = null, bool? KeepLines = null, bool? KeepWithNext = null,
    bool? PageBreakBefore = null, short? LineValue = null, bool? LineIsMultiple = null,
    IReadOnlyList<short>? ClearedTabPositions = null,
    IReadOnlyList<DocTabStop>? TabStops = null, uint? FillRgb = null,
    uint? ShadingForegroundRgb = null, ushort? ShadingPattern = null,
    DocParagraphBorder? TopBorder = null, DocParagraphBorder? LeftBorder = null,
    DocParagraphBorder? BottomBorder = null, DocParagraphBorder? RightBorder = null,
    DocParagraphBorder? BetweenBorder = null, bool? InTable = null,
    bool? TableTerminator = null, IReadOnlyList<short>? TableCellEdges = null,
    bool? TableHeader = null, bool? TableCantSplit = null,
    short? TableRowHeightTwips = null,
    IReadOnlyList<DocCellShading?>? TableCellShadings = null,
    IReadOnlyList<byte?>? TableCellVerticalAlignments = null,
    IReadOnlyList<DocCellBorders?>? TableCellBorders = null,
    DocTableBorders? TableBorders = null,
    short? ListOverrideIndex = null, byte? ListLevel = null,
    IReadOnlyList<byte?>? TableCellVerticalMerges = null,
    IReadOnlyList<byte?>? TableCellHorizontalMerges = null,
    bool? TableAutoFit = null, DocTablePreferredWidth? TablePreferredWidth = null,
    IReadOnlyList<DocTablePreferredWidth?>? TableCellPreferredWidths = null,
    DocCellMargins? TableDefaultCellMargins = null,
    IReadOnlyList<DocCellMargins?>? TableCellMargins = null,
    bool? TableRightToLeft = null, ushort? TableCellSpacingTwips = null,
    ushort? TableStyleIndex = null, DocCellShading? TableStyleShading = null,
    uint? ParagraphGroupId = null, uint? TableGroupId = null,
    IReadOnlyList<DocTabClearRange>? TabClearRanges = null,
    bool? WidowControl = null, short? TableIndentTwips = null,
    byte? TableJustification = null, bool? ContextualSpacing = null,
    int? TableDepth = null, bool? InnerTableCell = null,
    bool? InnerTableRow = null,
    IReadOnlyList<bool?>? TableCellNoWraps = null,
    IReadOnlyList<bool?>? TableCellFitTexts = null,
    DocTablePreferredWidth? TableWidthBefore = null,
    DocTablePreferredWidth? TableWidthAfter = null,
    short? TableRowOriginTwips = null,
    bool? ParagraphRightToLeft = null,
    bool? BeforeAutoSpacing = null, bool? AfterAutoSpacing = null,
    short? BeforeLines = null, short? AfterLines = null,
    short? LeftChars = null, short? RightChars = null,
    short? FirstLineChars = null,
    byte? TableHorizontalBandSize = null, byte? TableVerticalBandSize = null,
    ushort? TableLookMask = null, bool? MirrorIndents = null,
    bool? SuppressAutoHyphens = null,
    short? TextAlignmentCode = null,
    bool? SuppressLineNumbers = null,
    byte? TableStyleVerticalAlignment = null,
    IReadOnlyList<ushort?>? TableCellTextFlows = null,
    IReadOnlyList<bool?>? TableCellHideMarks = null,
    bool? TableStyleNoWrap = null, byte? OutlineLevel = null,
    bool? Kinsoku = null, bool? WordWrap = null,
    bool? SnapToGrid = null, bool? AutoSpaceDE = null,
    bool? AutoSpaceDN = null, bool? AdjustRightIndent = null,
    DocCellShading? TableBackgroundShading = null)
{
    public static DocParagraphFormatting Empty { get; } = new(null, null, null);
    public bool IsEmpty => Justification == null && BeforeTwips == null &&
        AfterTwips == null && LeftTwips == null && RightTwips == null &&
        FirstLineTwips == null && KeepLines == null && KeepWithNext == null &&
        PageBreakBefore == null && LineValue == null &&
        (ClearedTabPositions == null || ClearedTabPositions.Count == 0) &&
        (TabClearRanges == null || TabClearRanges.Count == 0) &&
        (TabStops == null || TabStops.Count == 0) && FillRgb == null &&
        ShadingForegroundRgb == null && ShadingPattern == null &&
        TopBorder == null && LeftBorder == null && BottomBorder == null &&
        RightBorder == null && BetweenBorder == null && InTable == null &&
        TableTerminator == null && TableCellEdges == null &&
        TableHeader == null && TableCantSplit == null && TableRowHeightTwips == null &&
        TableCellShadings == null && TableCellVerticalAlignments == null &&
        TableCellBorders == null && TableBorders == null &&
        ListOverrideIndex == null && ListLevel == null &&
        TableCellVerticalMerges == null && TableCellHorizontalMerges == null &&
        TableAutoFit == null && TablePreferredWidth == null &&
        TableCellPreferredWidths == null && TableCellNoWraps == null &&
        TableCellFitTexts == null &&
        TableDefaultCellMargins == null &&
        TableCellMargins == null && TableRightToLeft == null &&
        TableCellSpacingTwips == null && TableStyleIndex == null &&
        TableStyleShading == null && TableBackgroundShading == null && ParagraphGroupId == null && TableGroupId == null &&
        WidowControl == null && TableIndentTwips == null && TableJustification == null &&
        ContextualSpacing == null && TableDepth == null &&
        InnerTableCell == null && InnerTableRow == null &&
        TableWidthBefore == null && TableWidthAfter == null &&
        TableRowOriginTwips == null && ParagraphRightToLeft == null &&
        OutlineLevel == null && Kinsoku == null && WordWrap == null &&
        SnapToGrid == null && AutoSpaceDE == null && AutoSpaceDN == null &&
        AdjustRightIndent == null &&
        BeforeAutoSpacing == null && AfterAutoSpacing == null &&
        BeforeLines == null && AfterLines == null && LeftChars == null &&
        RightChars == null && FirstLineChars == null &&
        TableHorizontalBandSize == null && TableVerticalBandSize == null &&
        TableLookMask == null && MirrorIndents == null &&
        SuppressAutoHyphens == null && TextAlignmentCode == null &&
        SuppressLineNumbers == null && TableStyleVerticalAlignment == null &&
        TableStyleNoWrap == null &&
        TableCellTextFlows == null && TableCellHideMarks == null;

    internal static DocParagraphFormatting Parse(ReadOnlySpan<byte> sprms)
    {
        byte? justification = null;
        ushort? before = null, after = null;
        short? left = null, right = null, firstLine = null;
        bool? keepLines = null, keepWithNext = null, pageBreakBefore = null,
            widowControl = null, paragraphRightToLeft = null,
            kinsoku = null, wordWrap = null, snapToGrid = null,
            autoSpaceDE = null, autoSpaceDN = null, adjustRightIndent = null;
        short? lineValue = null;
        byte? outlineLevel = null;
        bool? lineIsMultiple = null;
        var clearedTabs = new List<short>();
        var tabClearRanges = new List<DocTabClearRange>();
        var tabStops = new List<DocTabStop>();
        uint? fillRgb = null, foregroundRgb = null;
        ushort? shadingPattern = null;
        var modernShadingSeen = false;
        static uint? LegacyShadingColor(int index) => index switch
        {
            0 => null,
            1 => 0x000000u, 2 => 0xFF0000u, 3 => 0xFFFF00u,
            4 => 0x00FF00u, 5 => 0xFF00FFu, 6 => 0x0000FFu,
            7 => 0x00FFFFu, 8 => 0xFFFFFFu, 9 => 0x800000u,
            10 => 0x808000u, 11 => 0x008000u,
            12 or 13 => 0x800080u, 14 => 0x008080u,
            15 => 0x808080u, 16 => 0xC0C0C0u,
            _ => throw new InvalidDataException("A legacy paragraph shading color is invalid.")
        };
        DocParagraphBorder? topBorder = null, leftBorder = null, bottomBorder = null,
            rightBorder = null, betweenBorder = null;
        bool? inTable = null, tableTerminator = null;
        int? tableDepth = null;
        bool? innerTableCell = null, innerTableRow = null;
        bool? tableHeader = null, tableCantSplit = null, contextualSpacing = null,
            mirrorIndents = null, suppressAutoHyphens = null;
        short? textAlignmentCode = null;
        bool? suppressLineNumbers = null;
        bool? beforeAutoSpacing = null, afterAutoSpacing = null;
        short? beforeLines = null, afterLines = null;
        short? leftChars = null, rightChars = null, firstLineChars = null;
        bool? tableAutoFit = null, tableRightToLeft = null;
        ushort? tableStyleIndex = null;
        short? tableIndentTwips = null;
        byte? tableJustification = null;
        byte? physicalTableJustification = null, logicalTableJustification = null;
        uint? paragraphGroupId = null, tableGroupId = null;
        ushort? tableCellSpacing = null;
        DocTablePreferredWidth? tablePreferredWidth = null;
        DocTablePreferredWidth? tableWidthBefore = null, tableWidthAfter = null;
        short? tableRowOriginTwips = null;
        var tableCellPreferredWidths = new List<DocTablePreferredWidth?>();
        var tableCellNoWraps = new List<bool?>();
        var tableCellFitTexts = new List<bool?>();
        DocCellMargins? tableDefaultCellMargins = null;
        var tableCellMargins = new List<DocCellMargins?>();
        short? tableRowHeight = null;
        IReadOnlyList<short>? tableCellEdges = null;
        var insertedCellWidths = new List<ushort>();
        var tableCellShadings = new List<DocCellShading?>();
        DocCellShading? tableBackgroundShading = null;
        var modernShadingCells = new HashSet<int>();
        void SetCellShading(int cell, DocCellShading? value, bool modern)
        {
            while (tableCellShadings.Count <= cell) tableCellShadings.Add(null);
            if (modern)
            {
                tableCellShadings[cell] = value;
                modernShadingCells.Add(cell);
            }
            else if (!modernShadingCells.Contains(cell))
                tableCellShadings[cell] = value;
        }
        DocCellShading? tableStyleShading = null;
        byte? tableStyleVerticalAlignment = null;
        bool? tableStyleNoWrap = null;
        byte? tableHorizontalBandSize = null, tableVerticalBandSize = null;
        ushort? tableLookMask = null;
        var tableCellVerticalAlignments = new List<byte?>();
        var tableCellTextFlows = new List<ushort?>();
        var tableCellHideMarks = new List<bool?>();
        var tableCellVerticalMerges = new List<byte?>();
        var tableCellHorizontalMerges = new List<byte?>();
        var tableCellBorders = new List<DocCellBorders?>();
        var tableCellBorderColors = new Dictionary<(int Cell, ushort Code), uint>();
        DocTableBorders? tableBorders = null;
        short? listOverrideIndex = null;
        byte? listLevel = null;
        for (var offset = 0; offset < sprms.Length;)
        {
            if (offset + 2 > sprms.Length)
                throw new InvalidDataException("A paragraph SPRM is truncated.");
            var sprm = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
            offset += 2;
            var spra = sprm >> 13;
            var length = spra switch
            {
                0 or 1 => 1,
                2 or 4 or 5 => 2,
                3 => 4,
                6 when sprm == 0xD608 && offset + 2 <= sprms.Length =>
                    1 + BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)),
                6 when sprm == 0xC615 && offset < sprms.Length &&
                    sprms[offset] == 255 => GetExtendedTabOperandLength(sprms.Slice(offset)),
                6 when offset < sprms.Length => 1 + sprms[offset],
                7 => 3,
                _ => throw new InvalidDataException("A paragraph SPRM has an invalid operand.")
            };
            if (offset + length > sprms.Length)
                throw new InvalidDataException("A paragraph SPRM operand is truncated.");
            switch (sprm)
            {
                case 0x347C:
                    if (sprms[offset] > 2)
                        throw new InvalidDataException("A table-style vertical alignment is invalid.");
                    tableStyleVerticalAlignment = sprms[offset];
                    break;
                case 0x347D:
                    if (sprms[offset] > 1)
                        throw new InvalidDataException("A table-style no-wrap value is invalid.");
                    tableStyleNoWrap = sprms[offset] != 0;
                    break;
                case 0x3488:
                    if (sprms[offset] is < 1 or > 3)
                        throw new InvalidDataException("A DOC row band size is invalid.");
                    tableHorizontalBandSize = sprms[offset];
                    break;
                case 0x3489:
                    if (sprms[offset] is < 1 or > 3)
                        throw new InvalidDataException("A DOC column band size is invalid.");
                    tableVerticalBandSize = sprms[offset];
                    break;
                case 0x740A:
                    if (length != 4)
                        throw new InvalidDataException("A DOC table look operand is invalid.");
                    tableLookMask = (ushort)(BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 2)) & 0x07E0);
                    break;
                case 0x2461: justification = sprms[offset]; break;
                case 0x260A: listLevel = sprms[offset]; break;
                case 0x460B:
                    listOverrideIndex = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    break;
                case 0xA413: before = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0xA414: after = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4458: beforeLines = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4459: afterLines = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4455: rightChars = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4456: leftChars = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x4457: firstLineChars = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x845D: right = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x845E: left = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x8460: firstLine = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset)); break;
                case 0x2405: keepLines = sprms[offset] != 0; break;
                case 0x2406: keepWithNext = sprms[offset] != 0; break;
                case 0x2407: pageBreakBefore = sprms[offset] != 0; break;
                case 0x2431: widowControl = sprms[offset] != 0; break;
                case 0x2433: kinsoku = sprms[offset] != 0; break;
                case 0x2434: wordWrap = sprms[offset] != 0; break;
                case 0x2437: autoSpaceDE = sprms[offset] != 0; break;
                case 0x2438: autoSpaceDN = sprms[offset] != 0; break;
                case 0x2448: adjustRightIndent = sprms[offset] != 0; break;
                case 0x2447: snapToGrid = sprms[offset] != 0; break;
                case 0x2640:
                    if (sprms[offset] > 9)
                        throw new InvalidDataException("A DOC outline level is invalid.");
                    outlineLevel = sprms[offset];
                    break;
                case 0x2441: paragraphRightToLeft = sprms[offset] != 0; break;
                case 0x246D: contextualSpacing = sprms[offset] != 0; break;
                case 0x2470: mirrorIndents = sprms[offset] != 0; break;
                case 0x242A: suppressAutoHyphens = sprms[offset] != 0; break;
                case 0x240C: suppressLineNumbers = sprms[offset] != 0; break;
                case 0x4439:
                    var alignment = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    if (alignment is >= 0 and <= 4) textAlignmentCode = alignment;
                    break;
                case 0x245B: beforeAutoSpacing = sprms[offset] != 0; break;
                case 0x245C: afterAutoSpacing = sprms[offset] != 0; break;
                case 0x2416: inTable = sprms[offset] != 0; break;
                case 0x2417: tableTerminator = sprms[offset] != 0; break;
                case 0x6649:
                    tableDepth = BinaryPrimitives.ReadInt32LittleEndian(sprms.Slice(offset));
                    if (tableDepth is < 0 or > 63)
                        throw new InvalidDataException("A DOC table depth is invalid.");
                    break;
                case 0x664A:
                    var depthDelta = BinaryPrimitives.ReadInt32LittleEndian(sprms.Slice(offset));
                    tableDepth = checked((tableDepth ?? (inTable == true ? 1 : 0)) + depthDelta);
                    if (tableDepth is < 0 or > 63)
                        throw new InvalidDataException("A DOC table depth is invalid.");
                    break;
                case 0x244B: innerTableCell = sprms[offset] != 0; break;
                case 0x244C: innerTableRow = sprms[offset] != 0; break;
                case 0x3404: tableHeader = sprms[offset] != 0; break;
                case 0x3615: tableAutoFit = sprms[offset] != 0; break;
                case 0x563A:
                    tableStyleIndex = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset));
                    break;
                case 0x6465:
                    paragraphGroupId = BinaryPrimitives.ReadUInt32LittleEndian(
                        sprms.Slice(offset));
                    break;
                case 0x7469:
                    tableGroupId = BinaryPrimitives.ReadUInt32LittleEndian(
                        sprms.Slice(offset));
                    break;
                case 0x560B or 0x5664:
                    if (length != 2 || BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset)) > 1)
                        throw new InvalidDataException("A DOC table has an invalid direction flag.");
                    tableRightToLeft = tableRightToLeft == true || sprms[offset] != 0;
                    break;
                case 0xF614:
                    if (length != 3 || sprms[offset] is not (0 or 1 or 2 or 3))
                        throw new InvalidDataException("A DOC table has invalid preferred-width units.");
                    if (sprms[offset] == 0) break;
                    var preferred = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 1));
                    if ((sprms[offset] == 1 && preferred != 0) ||
                        (sprms[offset] == 2 && preferred > 30000) ||
                        (sprms[offset] == 3 && preferred > 31680))
                        throw new InvalidDataException("A DOC table has an invalid preferred width.");
                    tablePreferredWidth = new DocTablePreferredWidth(sprms[offset], preferred);
                    break;
                case 0xF617 or 0xF618:
                    if (length != 3 || sprms[offset] is not (0 or 1 or 2 or 3))
                        throw new InvalidDataException("A DOC row gap has invalid width units.");
                    // ftsNil leaves this row gap unspecified, including its width value.
                    if (sprms[offset] == 0) break;
                    var gapWidth = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 1));
                    if ((sprms[offset] == 1 && gapWidth != 0) ||
                        (sprms[offset] == 2 && gapWidth > 5000) ||
                        (sprms[offset] == 3 && gapWidth > 31680))
                        throw new InvalidDataException("A DOC row gap has an invalid width.");
                    var gap = new DocTablePreferredWidth(sprms[offset], gapWidth);
                    if (sprm == 0xF617) tableWidthBefore = gap;
                    else tableWidthAfter = gap;
                    break;
                case 0x9601:
                    tableRowOriginTwips = BinaryPrimitives.ReadInt16LittleEndian(
                        sprms.Slice(offset));
                    break;
                case 0xF661:
                    if (length != 3 || sprms[offset] is not (0 or 1 or 3))
                        throw new InvalidDataException("A DOC table has invalid indent units.");
                    var indent = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset + 1));
                    if (sprms[offset] != 3 && indent != 0)
                        throw new InvalidDataException("A DOC table has an invalid automatic indent.");
                    if (indent is < -31560 or > 31680)
                        throw new InvalidDataException("A DOC table indent exceeds the DOC limit.");
                    tableIndentTwips = sprms[offset] == 3 ? indent : null;
                    break;
                case 0x5400 or 0x548A:
                    var tableAlignment = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset));
                    if (tableAlignment > 2)
                        throw new InvalidDataException("A DOC table has invalid justification.");
                    if (sprm == 0x548A)
                        logicalTableJustification = (byte)tableAlignment;
                    else physicalTableJustification = (byte)tableAlignment;
                    break;
                case 0xD635:
                    if (length != 6 || sprms[offset] != 5 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63 ||
                        sprms[offset + 3] is not (0 or 1 or 2 or 3))
                        throw new InvalidDataException("A DOC cell has invalid preferred-width units.");
                    if (sprms[offset + 3] == 0) break;
                    var cellPreferred = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 4));
                    if ((sprms[offset + 3] == 1 && cellPreferred != 0) ||
                        (sprms[offset + 3] == 2 && cellPreferred > 5000) ||
                        (sprms[offset + 3] == 3 && cellPreferred > 31680))
                        throw new InvalidDataException("A DOC cell has an invalid preferred width.");
                    while (tableCellPreferredWidths.Count < sprms[offset + 2])
                        tableCellPreferredWidths.Add(null);
                    for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                        tableCellPreferredWidths[i] = new DocTablePreferredWidth(
                            sprms[offset + 3], cellPreferred);
                    break;
                case 0x7621:
                    if (length != 4 || sprms[offset] > insertedCellWidths.Count ||
                        sprms[offset] >= sprms[offset + 1] ||
                        sprms[offset + 1] > 63 ||
                        insertedCellWidths.Count + sprms[offset + 1] - sprms[offset] > 63)
                        throw new InvalidDataException("A DOC table cell insertion is invalid.");
                    var insertedWidth = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 2));
                    if (insertedWidth > 31680)
                        throw new InvalidDataException("A DOC table cell insertion is too wide.");
                    insertedCellWidths.InsertRange(sprms[offset], Enumerable.Repeat(
                        insertedWidth, sprms[offset + 1] - sprms[offset]));
                    break;
                case 0x7623:
                    if (length != 4 || sprms[offset] >= sprms[offset + 1] ||
                        sprms[offset + 1] > insertedCellWidths.Count)
                        throw new InvalidDataException("A DOC table column range is invalid.");
                    var columnWidth = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 2));
                    if (columnWidth > 31680)
                        throw new InvalidDataException("A DOC table column is too wide.");
                    for (var i = sprms[offset]; i < sprms[offset + 1]; i++)
                        insertedCellWidths[i] = columnWidth;
                    break;
                case 0x5624:
                    if (length != 2 || sprms[offset] + 1 >= sprms[offset + 1] ||
                        sprms[offset + 1] > 63 ||
                        (insertedCellWidths.Count != 0 &&
                         sprms[offset + 1] > insertedCellWidths.Count))
                        throw new InvalidDataException("A DOC table merge range is invalid.");
                    while (tableCellHorizontalMerges.Count < sprms[offset + 1])
                        tableCellHorizontalMerges.Add(null);
                    for (var i = sprms[offset]; i < sprms[offset + 1]; i++)
                    {
                        if (tableCellHorizontalMerges[i] != null)
                            throw new InvalidDataException("DOC table merge ranges overlap.");
                        tableCellHorizontalMerges[i] = i == sprms[offset] ?
                            (byte)2 : (byte)1;
                    }
                    break;
                case 0xD639:
                    if (length != 4 || sprms[offset] != 3 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63 || sprms[offset + 3] > 1)
                        throw new InvalidDataException("A DOC cell no-wrap operand is invalid.");
                    while (tableCellNoWraps.Count < sprms[offset + 2])
                        tableCellNoWraps.Add(null);
                    for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                        tableCellNoWraps[i] = sprms[offset + 3] != 0;
                    break;
                case 0xF636:
                    if (length != 3 || sprms[offset] >= sprms[offset + 1] ||
                        sprms[offset + 1] > 63 || sprms[offset + 2] > 1)
                        throw new InvalidDataException("A DOC cell fit-text operand is invalid.");
                    while (tableCellFitTexts.Count < sprms[offset + 1])
                        tableCellFitTexts.Add(null);
                    for (var i = sprms[offset]; i < sprms[offset + 1]; i++)
                        tableCellFitTexts[i] = sprms[offset + 2] != 0;
                    break;
                case 0xD632 or 0xD634:
                    if (length != 7 || sprms[offset] != 6 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63 ||
                        (sprms[offset + 3] & 0xF0) != 0 ||
                        sprms[offset + 4] is not (0 or 3))
                        throw new InvalidDataException("A DOC cell margin operand is invalid.");
                    if (sprms[offset + 4] == 0) break;
                    var marginWidth = BinaryPrimitives.ReadUInt16LittleEndian(
                        sprms.Slice(offset + 5));
                    if (marginWidth > 31680)
                        throw new InvalidDataException("A DOC cell margin exceeds 22 inches.");
                    var sides = sprms[offset + 3];
                    DocCellMargins ApplyMargin(DocCellMargins? current)
                    {
                        current ??= new DocCellMargins();
                        return current with
                        {
                            Top = (sides & 1) != 0 ? marginWidth : current.Top,
                            Left = (sides & 2) != 0 ? marginWidth : current.Left,
                            Bottom = (sides & 4) != 0 ? marginWidth : current.Bottom,
                            Right = (sides & 8) != 0 ? marginWidth : current.Right
                        };
                    }
                    if (sprm == 0xD634)
                    {
                        if (sprms[offset + 1] != 0 || sprms[offset + 2] != 1)
                            throw new InvalidDataException("A DOC default cell margin range is invalid.");
                        tableDefaultCellMargins = ApplyMargin(tableDefaultCellMargins);
                    }
                    else
                    {
                        while (tableCellMargins.Count < sprms[offset + 2])
                            tableCellMargins.Add(null);
                        for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                            tableCellMargins[i] = ApplyMargin(tableCellMargins[i]);
                    }
                    break;
                case 0xD633:
                    if (length != 7 || sprms[offset] != 6 ||
                        sprms[offset + 1] != 0 || sprms[offset + 2] != 1 ||
                        sprms[offset + 3] != 0x0F ||
                        sprms[offset + 4] is not (0 or 3 or 0x13))
                        throw new InvalidDataException("A DOC cell spacing operand is invalid.");
                    tableCellSpacing = sprms[offset + 4] == 0 ? (ushort)0 :
                        BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset + 5));
                    if (tableCellSpacing > 15840)
                        throw new InvalidDataException("A DOC cell spacing exceeds 11 inches.");
                    break;
                case 0x3403 or 0x3466: tableCantSplit = sprms[offset] != 0; break;
                case 0x9407:
                    tableRowHeight = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    break;
                case 0xD608:
                    if (length < 5 || sprms[offset] + (sprms[offset + 1] << 8) + 1 != length)
                        throw new InvalidDataException("A table row definition has an invalid length.");
                    var columnCount = sprms[offset + 2];
                    if (columnCount is < 1 or > 63 || length < 3 + (columnCount + 1) * 2)
                        throw new InvalidDataException("A table row definition has invalid columns.");
                    var edges = new short[columnCount + 1];
                    for (var i = 0; i < edges.Length; i++)
                        edges[i] = BinaryPrimitives.ReadInt16LittleEndian(
                            sprms.Slice(offset + 3 + i * 2));
                    tableCellEdges = edges;
                    var cellsOffset = offset + 3 + (columnCount + 1) * 2;
                    var availableCells = Math.Min(columnCount,
                        (offset + length - cellsOffset) / 20);
                    for (var i = 0; i < availableCells; i++)
                    {
                        var flags = BinaryPrimitives.ReadUInt16LittleEndian(
                            sprms.Slice(cellsOffset + i * 20, 2));
                        var horizontal = (byte)(flags & 3);
                        var merge = (byte)((flags >> 5) & 3);
                        var verticalAlignment = (byte)((flags >> 7) & 3);
                        var textFlow = (ushort)((flags >> 2) & 7);
                        while (tableCellTextFlows.Count <= i)
                            tableCellTextFlows.Add(null);
                        if (textFlow is 1 or 3 or 4 or 5)
                            tableCellTextFlows[i] = textFlow;
                        while (tableCellFitTexts.Count <= i)
                            tableCellFitTexts.Add(null);
                        while (tableCellNoWraps.Count <= i)
                            tableCellNoWraps.Add(null);
                        while (tableCellHideMarks.Count <= i)
                            tableCellHideMarks.Add(null);
                        if ((flags & 0x1000) != 0) tableCellFitTexts[i] = true;
                        if ((flags & 0x2000) != 0) tableCellNoWraps[i] = true;
                        if ((flags & 0x4000) != 0) tableCellHideMarks[i] = true;
                        tableCellHorizontalMerges.Add(horizontal == 0 ? null : horizontal);
                        tableCellVerticalMerges.Add(merge == 0 ? null : merge);
                        while (tableCellVerticalAlignments.Count <= i)
                            tableCellVerticalAlignments.Add(null);
                        if (verticalAlignment is 1 or 2)
                            tableCellVerticalAlignments[i] = verticalAlignment;
                        var cell = sprms.Slice(cellsOffset + i * 20, 20);
                        var borders = new DocCellBorders(
                            ParseOptionalCellBorder80(cell.Slice(4, 4)),
                            ParseOptionalCellBorder80(cell.Slice(8, 4)),
                            ParseOptionalCellBorder80(cell.Slice(12, 4)),
                            ParseOptionalCellBorder80(cell.Slice(16, 4)));
                        while (tableCellBorders.Count <= i)
                            tableCellBorders.Add(null);
                        if (borders.Top != null || borders.Left != null ||
                            borders.Bottom != null || borders.Right != null)
                            tableCellBorders[i] = borders;
                    }
                    break;
                case 0xD609:
                {
                    if (length < 1 || sprms[offset] + 1 != length ||
                        sprms[offset] % 2 != 0 || sprms[offset] > 126)
                        throw new InvalidDataException("A legacy table shading operand has an invalid length.");
                    for (var i = 0; i < sprms[offset] / 2; i++)
                    {
                        var cellShd80 = BinaryPrimitives.ReadUInt16LittleEndian(
                            sprms.Slice(offset + 1 + i * 2));
                        var cellForegroundIndex = cellShd80 & 31;
                        var cellBackgroundIndex = (cellShd80 >> 5) & 31;
                        var cellPattern = (ushort)(cellShd80 >> 10);
                        if ((cellForegroundIndex == 31 && cellBackgroundIndex == 31 &&
                            cellPattern == 63) ||
                            (cellForegroundIndex == 0 && cellBackgroundIndex == 0 &&
                            cellPattern == 0))
                        {
                            SetCellShading(i, null, false);
                            continue;
                        }
                        if (DocShadingPatterns.ToOpenXml(cellPattern) == null)
                            throw new InvalidDataException("A legacy table shading pattern is invalid.");
                        SetCellShading(i, new DocCellShading(
                            LegacyShadingColor(cellBackgroundIndex),
                            LegacyShadingColor(cellForegroundIndex), cellPattern), false);
                    }
                    break;
                }
                case 0xD670 or 0xD671 or 0xD672:
                    if (length < 1 || sprms[offset] + 1 != length ||
                        sprms[offset] % 10 != 0)
                        throw new InvalidDataException("A table shading operand has an invalid length.");
                    var firstCell = (sprm - 0xD670) * 22;
                    for (var i = 0; i < sprms[offset] / 10; i++)
                    {
                        var shade = sprms.Slice(offset + 1 + i * 10, 10);
                        var foregroundColor = BinaryPrimitives.ReadUInt32LittleEndian(shade);
                        var backgroundColor = BinaryPrimitives.ReadUInt32LittleEndian(shade.Slice(4));
                        var shadePattern = BinaryPrimitives.ReadUInt16LittleEndian(shade.Slice(8));
                        SetCellShading(firstCell + i,
                            foregroundColor == 0xFF000000 &&
                            backgroundColor == 0xFF000000 && shadePattern == 0
                            ? null : foregroundColor == 0xFFFFFFFF &&
                            backgroundColor == 0xFF000000 && shadePattern == 0
                            ? new DocCellShading(null, null, ushort.MaxValue)
                            : DocShadingPatterns.ToOpenXml(shadePattern) == null
                            ? null : new DocCellShading(
                                (backgroundColor >> 24) == 0 ? backgroundColor & 0xFFFFFF : null,
                                (foregroundColor >> 24) == 0 ? foregroundColor & 0xFFFFFF : null,
                                shadePattern), true);
                    }
                    break;
                case 0xD660:
                    if (length != 11 || sprms[offset] != 10)
                        throw new InvalidDataException("A table background shading operand is invalid.");
                    var backgroundShade = sprms.Slice(offset + 1, 10);
                    var backgroundForeground = BinaryPrimitives.ReadUInt32LittleEndian(backgroundShade);
                    var backgroundFill = BinaryPrimitives.ReadUInt32LittleEndian(backgroundShade.Slice(4));
                    var backgroundPattern = BinaryPrimitives.ReadUInt16LittleEndian(backgroundShade.Slice(8));
                    if (DocShadingPatterns.ToOpenXml(backgroundPattern) != null)
                        tableBackgroundShading = new DocCellShading(
                            backgroundFill >> 24 == 0 ? backgroundFill & 0xFFFFFF : null,
                            backgroundForeground >> 24 == 0 ? backgroundForeground & 0xFFFFFF : null,
                            backgroundPattern);
                    break;
                case 0xD687:
                    if (length != 11 || sprms[offset] != 10)
                        throw new InvalidDataException("A table-style shading operand is invalid.");
                    var styleShade = sprms.Slice(offset + 1, 10);
                    var styleForeground = BinaryPrimitives.ReadUInt32LittleEndian(styleShade);
                    var styleBackground = BinaryPrimitives.ReadUInt32LittleEndian(styleShade.Slice(4));
                    var stylePattern = BinaryPrimitives.ReadUInt16LittleEndian(styleShade.Slice(8));
                    if (DocShadingPatterns.ToOpenXml(stylePattern) != null)
                        tableStyleShading = new DocCellShading(
                            styleBackground >> 24 == 0 ? styleBackground & 0xFFFFFF : null,
                            styleForeground >> 24 == 0 ? styleForeground & 0xFFFFFF : null,
                            stylePattern);
                    break;
                case 0xD62C:
                    if (length != 4 || sprms[offset] != 3 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63 || sprms[offset + 3] > 2)
                        throw new InvalidDataException("A table cell vertical alignment is invalid.");
                    while (tableCellVerticalAlignments.Count < sprms[offset + 2])
                        tableCellVerticalAlignments.Add(null);
                    for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                        tableCellVerticalAlignments[i] = sprms[offset + 3];
                    break;
                case 0xD642:
                    if (length != 4 || sprms[offset] != 3 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63 || sprms[offset + 3] > 1)
                        throw new InvalidDataException("A table cell hide-mark operand is invalid.");
                    while (tableCellHideMarks.Count < sprms[offset + 2])
                        tableCellHideMarks.Add(null);
                    for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                        tableCellHideMarks[i] = sprms[offset + 3] != 0;
                    break;
                case 0x7629:
                    if (length != 4 || sprms[offset] >= sprms[offset + 1] ||
                        sprms[offset + 1] > 63)
                        throw new InvalidDataException("A table cell text flow range is invalid.");
                    var flow = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset + 2));
                    if (flow is not (0 or 1 or 3 or 4 or 5))
                        throw new InvalidDataException("A table cell text flow is invalid.");
                    while (tableCellTextFlows.Count < sprms[offset + 1])
                        tableCellTextFlows.Add(null);
                    for (var i = sprms[offset]; i < sprms[offset + 1]; i++)
                        tableCellTextFlows[i] = flow;
                    break;
                case 0xD62F:
                    if (length != 12 || sprms[offset] != 11 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63)
                        throw new InvalidDataException("A table cell border operand is invalid.");
                    var border = DocParagraphBorder.ParseRaw(sprms.Slice(offset + 4, 8));
                    if (border == null) break;
                    while (tableCellBorders.Count < sprms[offset + 2])
                        tableCellBorders.Add(null);
                    for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                    {
                        var current = tableCellBorders[i] ?? new DocCellBorders();
                        if ((sprms[offset + 3] & 1) != 0) current = current with { Top = border };
                        if ((sprms[offset + 3] & 2) != 0) current = current with { Left = border };
                        if ((sprms[offset + 3] & 4) != 0) current = current with { Bottom = border };
                        if ((sprms[offset + 3] & 8) != 0) current = current with { Right = border };
                        if ((sprms[offset + 3] & 0x10) != 0)
                            current = current with { TopLeftToBottomRight = border };
                        if ((sprms[offset + 3] & 0x20) != 0)
                            current = current with { TopRightToBottomLeft = border };
                        tableCellBorders[i] = current;
                    }
                    break;
                case 0xD620:
                    if (length != 8 || sprms[offset] != 7 ||
                        sprms[offset + 1] >= sprms[offset + 2] ||
                        sprms[offset + 2] > 63)
                        throw new InvalidDataException("A legacy table cell border operand is invalid.");
                    var legacyCellBorder = DocParagraphBorder.Parse80(
                        sprms.Slice(offset + 4, 4));
                    if (legacyCellBorder == null) break;
                    while (tableCellBorders.Count < sprms[offset + 2])
                        tableCellBorders.Add(null);
                    for (var i = sprms[offset + 1]; i < sprms[offset + 2]; i++)
                    {
                        var current = tableCellBorders[i] ?? new DocCellBorders();
                        if ((sprms[offset + 3] & 1) != 0)
                            current = current with { Top = legacyCellBorder };
                        if ((sprms[offset + 3] & 2) != 0)
                            current = current with { Left = legacyCellBorder };
                        if ((sprms[offset + 3] & 4) != 0)
                            current = current with { Bottom = legacyCellBorder };
                        if ((sprms[offset + 3] & 8) != 0)
                            current = current with { Right = legacyCellBorder };
                        if ((sprms[offset + 3] & 0x10) != 0)
                            current = current with { TopLeftToBottomRight = legacyCellBorder };
                        if ((sprms[offset + 3] & 0x20) != 0)
                            current = current with { TopRightToBottomLeft = legacyCellBorder };
                        tableCellBorders[i] = current;
                    }
                    break;
                case 0xD605:
                    if (length != 25 || sprms[offset] != 24)
                        throw new InvalidDataException("A legacy table border operand is invalid.");
                    tableBorders = new DocTableBorders(
                        DocParagraphBorder.Parse80(sprms.Slice(offset + 1, 4)),
                        DocParagraphBorder.Parse80(sprms.Slice(offset + 5, 4)),
                        DocParagraphBorder.Parse80(sprms.Slice(offset + 9, 4)),
                        DocParagraphBorder.Parse80(sprms.Slice(offset + 13, 4)),
                        DocParagraphBorder.Parse80(sprms.Slice(offset + 17, 4)),
                        DocParagraphBorder.Parse80(sprms.Slice(offset + 21, 4)));
                    break;
                case 0xD613:
                    if (length != 49 || sprms[offset] != 48)
                        throw new InvalidDataException("A table border operand is invalid.");
                    tableBorders = new DocTableBorders(
                        DocParagraphBorder.ParseRaw(sprms.Slice(offset + 1, 8)),
                        DocParagraphBorder.ParseRaw(sprms.Slice(offset + 9, 8)),
                        DocParagraphBorder.ParseRaw(sprms.Slice(offset + 17, 8)),
                        DocParagraphBorder.ParseRaw(sprms.Slice(offset + 25, 8)),
                        DocParagraphBorder.ParseRaw(sprms.Slice(offset + 33, 8)),
                        DocParagraphBorder.ParseRaw(sprms.Slice(offset + 41, 8)));
                    break;
                case 0xD61A or 0xD61B or 0xD61C or 0xD61D:
                    if (length < 1 || sprms[offset] != length - 1 ||
                        (length - 1) % 4 != 0 || (length - 1) / 4 > 63)
                        throw new InvalidDataException("A table border-color operand is invalid.");
                    for (var cell = 0; cell < (length - 1) / 4; cell++)
                        tableCellBorderColors[(cell, sprm)] =
                            BinaryPrimitives.ReadUInt32LittleEndian(
                                sprms.Slice(offset + 1 + cell * 4, 4));
                    break;
                case 0x6412:
                    lineValue = BinaryPrimitives.ReadInt16LittleEndian(sprms.Slice(offset));
                    var multiplier = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset + 2));
                    if (multiplier > 1 || (multiplier == 1 && lineValue < 0))
                        throw new InvalidDataException("A paragraph has invalid line spacing.");
                    lineIsMultiple = multiplier == 1;
                    break;
                case 0xC60D:
                case 0xC615:
                    ParseTabs(sprms.Slice(offset, length), clearedTabs, tabClearRanges, tabStops,
                        sprm == 0xC615);
                    break;
                case 0x442D when !modernShadingSeen:
                    var shd80 = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
                    var foregroundIndex = shd80 & 31;
                    var backgroundIndex = (shd80 >> 5) & 31;
                    var legacyPattern = (ushort)(shd80 >> 10);
                    if (foregroundIndex == 31 && backgroundIndex == 31 &&
                        legacyPattern == 63)
                    {
                        fillRgb = foregroundRgb = null;
                        shadingPattern = 0;
                        break;
                    }
                    if (DocShadingPatterns.ToOpenXml(legacyPattern) == null)
                        throw new InvalidDataException("A legacy paragraph shading pattern is invalid.");
                    foregroundRgb = LegacyShadingColor(foregroundIndex);
                    fillRgb = LegacyShadingColor(backgroundIndex);
                    shadingPattern = legacyPattern;
                    break;
                case 0xC64D when length == 11 && sprms[offset] == 10:
                    var pattern = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset + 9));
                    var foreground = BinaryPrimitives.ReadUInt32LittleEndian(sprms.Slice(offset + 1));
                    var background = BinaryPrimitives.ReadUInt32LittleEndian(sprms.Slice(offset + 5));
                    if (foreground == 0xFFFFFFFFu &&
                        background == 0xFF000000u && pattern == 0)
                    {
                        shadingPattern = ushort.MaxValue;
                        foregroundRgb = fillRgb = null;
                    }
                    else if (DocShadingPatterns.ToOpenXml(pattern) != null)
                    {
                        shadingPattern = pattern;
                        foregroundRgb = (foreground >> 24) == 0
                            ? foreground & 0xFFFFFF : null;
                        fillRgb = (background >> 24) == 0
                            ? background & 0xFFFFFF : null;
                    }
                    modernShadingSeen = true;
                    break;
                case 0x6424: topBorder = DocParagraphBorder.Parse80(sprms.Slice(offset, length)); break;
                case 0x6425: leftBorder = DocParagraphBorder.Parse80(sprms.Slice(offset, length)); break;
                case 0x6426: bottomBorder = DocParagraphBorder.Parse80(sprms.Slice(offset, length)); break;
                case 0x6427: rightBorder = DocParagraphBorder.Parse80(sprms.Slice(offset, length)); break;
                case 0x6428: betweenBorder = DocParagraphBorder.Parse80(sprms.Slice(offset, length)); break;
                case 0xC64E: topBorder = DocParagraphBorder.Parse(sprms.Slice(offset, length)); break;
                case 0xC64F: leftBorder = DocParagraphBorder.Parse(sprms.Slice(offset, length)); break;
                case 0xC650: bottomBorder = DocParagraphBorder.Parse(sprms.Slice(offset, length)); break;
                case 0xC651: rightBorder = DocParagraphBorder.Parse(sprms.Slice(offset, length)); break;
                case 0xC652: betweenBorder = DocParagraphBorder.Parse(sprms.Slice(offset, length)); break;
            }
            offset += length;
        }
        if (tableCellEdges == null && insertedCellWidths.Count > 0)
        {
            var edges = new short[insertedCellWidths.Count + 1];
            var position = 0;
            for (var i = 0; i < insertedCellWidths.Count; i++)
            {
                position += insertedCellWidths[i];
                if (position > 31680)
                    throw new InvalidDataException("A DOC table row is too wide.");
                edges[i + 1] = checked((short)position);
            }
            tableCellEdges = edges;
        }
        if (tableCellHorizontalMerges.Any(x => x != null) &&
            tableCellEdges is { Count: > 1 } rowEdges)
        {
            if (tableCellHorizontalMerges.Count > rowEdges.Count - 1)
                throw new InvalidDataException("A DOC table merge exceeds its cell count.");
            while (tableCellHorizontalMerges.Count < rowEdges.Count - 1)
                tableCellHorizontalMerges.Add(null);
        }
        foreach (var entry in tableCellBorderColors)
        {
            var cell = entry.Key.Cell;
            var code = entry.Key.Code;
            var rawColor = entry.Value;
            if (cell >= tableCellBorders.Count || tableCellBorders[cell] is not { } borders)
                continue;
            var border = code switch
            {
                0xD61A => borders.Top,
                0xD61B => borders.Left,
                0xD61C => borders.Bottom,
                _ => borders.Right
            };
            if (border == null && rawColor != uint.MaxValue) continue;
            var updated = rawColor == uint.MaxValue
                ? new DocParagraphBorder(0, 0, 0, null) : border! with
            {
                ColorRgb = (rawColor >> 24) == 0 ? rawColor & 0xFFFFFF : null
            };
            tableCellBorders[cell] = code switch
            {
                0xD61A => borders with { Top = updated },
                0xD61B => borders with { Left = updated },
                0xD61C => borders with { Bottom = updated },
                _ => borders with { Right = updated }
            };
        }
        tableJustification = logicalTableJustification ??
            (tableRightToLeft == true ? physicalTableJustification switch
            {
                0 => (byte)2,
                1 => (byte)1,
                2 => null,
                _ => null
            } : physicalTableJustification);
        return new DocParagraphFormatting(justification, before, after, left, right,
            firstLine, keepLines, keepWithNext, pageBreakBefore, lineValue,
            lineIsMultiple, clearedTabs, tabStops, fillRgb, foregroundRgb,
            shadingPattern, topBorder, leftBorder, bottomBorder, rightBorder,
            betweenBorder, inTable, tableTerminator, tableCellEdges,
            tableHeader, tableCantSplit, tableRowHeight,
            tableCellShadings.Count == 0 ? null : tableCellShadings,
            tableCellVerticalAlignments.Count == 0 ? null : tableCellVerticalAlignments,
            tableCellBorders.Count == 0 ? null : tableCellBorders, tableBorders,
            listOverrideIndex, listLevel,
            tableCellVerticalMerges.Any(x => x != null) ? tableCellVerticalMerges : null,
            tableCellHorizontalMerges.Any(x => x != null) ? tableCellHorizontalMerges : null,
            tableAutoFit, tablePreferredWidth,
            tableCellPreferredWidths.Count == 0 ? null : tableCellPreferredWidths,
            tableDefaultCellMargins,
            tableCellMargins.Count == 0 ? null : tableCellMargins,
            tableRightToLeft, tableCellSpacing, tableStyleIndex, tableStyleShading,
            paragraphGroupId, tableGroupId, tabClearRanges, widowControl,
            tableIndentTwips, tableJustification, contextualSpacing,
            tableDepth, innerTableCell, innerTableRow,
            tableCellNoWraps.Any(x => x != null) ? tableCellNoWraps : null,
            tableCellFitTexts.Any(x => x != null) ? tableCellFitTexts : null,
            tableWidthBefore, tableWidthAfter, tableRowOriginTwips,
            paragraphRightToLeft, beforeAutoSpacing, afterAutoSpacing,
            beforeLines, afterLines, leftChars, rightChars, firstLineChars,
            tableHorizontalBandSize, tableVerticalBandSize, tableLookMask,
            mirrorIndents, suppressAutoHyphens, textAlignmentCode,
            suppressLineNumbers, tableStyleVerticalAlignment,
            tableCellTextFlows.Any(x => x != null) ? tableCellTextFlows : null,
            tableCellHideMarks.Any(x => x != null) ? tableCellHideMarks : null,
            tableStyleNoWrap, outlineLevel, kinsoku, wordWrap, snapToGrid,
            autoSpaceDE, autoSpaceDN, adjustRightIndent, tableBackgroundShading);
    }

    internal byte[] Encode(bool forStyle = false)
    {
        using var stream = new MemoryStream();
        if (Justification is byte justification)
        {
            stream.WriteByte(0x61); stream.WriteByte(0x24); stream.WriteByte(justification);
        }
        if (ListLevel is byte listLevel)
        {
            stream.WriteByte(0x0A); stream.WriteByte(0x26); stream.WriteByte(listLevel);
        }
        if (ListOverrideIndex is short listOverrideIndex)
        {
            stream.WriteByte(0x0B); stream.WriteByte(0x46);
            WriteI16(stream, listOverrideIndex);
        }
        if (KeepLines is bool keepLines) WriteBool(stream, 0x05, keepLines);
        if (KeepWithNext is bool keepWithNext) WriteBool(stream, 0x06, keepWithNext);
        if (PageBreakBefore is bool pageBreakBefore) WriteBool(stream, 0x07, pageBreakBefore);
        if (WidowControl is bool widowControl) WriteBool(stream, 0x31, widowControl);
        if (Kinsoku is bool kinsoku) WriteBool(stream, 0x33, kinsoku);
        if (WordWrap is bool wordWrap) WriteBool(stream, 0x34, wordWrap);
        if (AutoSpaceDE is bool autoSpaceDE) WriteBool(stream, 0x37, autoSpaceDE);
        if (AutoSpaceDN is bool autoSpaceDN) WriteBool(stream, 0x38, autoSpaceDN);
        if (SnapToGrid is bool snapToGrid) WriteBool(stream, 0x47, snapToGrid);
        if (AdjustRightIndent is bool adjustRightIndent)
            WriteBool(stream, 0x48, adjustRightIndent);
        if (OutlineLevel is byte outlineLevel)
        {
            if (outlineLevel > 9)
                throw new InvalidDataException("A DOC outline level is invalid.");
            stream.WriteByte(0x40); stream.WriteByte(0x26);
            stream.WriteByte(outlineLevel);
        }
        if (ParagraphRightToLeft is bool paragraphRightToLeft)
            WriteBool(stream, 0x41, paragraphRightToLeft);
        if (ContextualSpacing is bool contextualSpacing)
            WriteBool(stream, 0x6D, contextualSpacing);
        if (MirrorIndents is bool mirrorIndents)
            WriteBool(stream, 0x70, mirrorIndents);
        if (SuppressAutoHyphens is bool suppressAutoHyphens)
            WriteBool(stream, 0x2A, suppressAutoHyphens);
        if (SuppressLineNumbers is bool suppressLineNumbers)
            WriteBool(stream, 0x0C, suppressLineNumbers);
        if (TextAlignmentCode is short textAlignmentCode)
        {
            stream.WriteByte(0x39); stream.WriteByte(0x44);
            WriteI16(stream, textAlignmentCode);
        }
        if (BeforeAutoSpacing is bool beforeAutoSpacing)
            WriteBool(stream, 0x5B, beforeAutoSpacing);
        if (AfterAutoSpacing is bool afterAutoSpacing)
            WriteBool(stream, 0x5C, afterAutoSpacing);
        if (InTable is bool inTable) WriteBool(stream, 0x16, inTable);
        if (TableTerminator is bool tableTerminator) WriteBool(stream, 0x17, tableTerminator);
        if (TableDepth is int tableDepth)
        {
            if (tableDepth is < 0 or > 63)
                throw new InvalidDataException("A DOC table depth is invalid.");
            stream.WriteByte(0x49); stream.WriteByte(0x66);
            WriteU32(stream, checked((uint)tableDepth));
        }
        if (InnerTableCell is bool innerTableCell) WriteBool(stream, 0x4B, innerTableCell);
        if (InnerTableRow is bool innerTableRow) WriteBool(stream, 0x4C, innerTableRow);
        if (TableHeader is bool tableHeader)
        {
            stream.WriteByte(0x04); stream.WriteByte(0x34);
            stream.WriteByte(tableHeader ? (byte)1 : (byte)0);
        }
        if (TableRightToLeft is bool tableRightToLeft)
        {
            stream.WriteByte(0x0B); stream.WriteByte(0x56);
            WriteI16(stream, tableRightToLeft ? (short)1 : (short)0);
        }
        if (TableStyleIndex is ushort tableStyleIndex)
        {
            stream.WriteByte(0x3A); stream.WriteByte(0x56);
            WriteI16(stream, unchecked((short)tableStyleIndex));
        }
        if (TableLookMask is ushort tableLookMask)
        {
            stream.WriteByte(0x0A); stream.WriteByte(0x74);
            WriteI16(stream, -1);
            WriteI16(stream, checked((short)(tableLookMask & 0x07E0)));
        }
        if (ParagraphGroupId is uint paragraphGroupId)
        {
            stream.WriteByte(0x65); stream.WriteByte(0x64);
            WriteU32(stream, paragraphGroupId);
        }
        if (TableGroupId is uint tableGroupId)
        {
            stream.WriteByte(0x69); stream.WriteByte(0x74);
            WriteU32(stream, tableGroupId);
        }
        if (TableBackgroundShading is { } tableBackgroundShading)
        {
            stream.WriteByte(0x60); stream.WriteByte(0xD6); stream.WriteByte(10);
            WriteColor(stream, tableBackgroundShading.ForegroundRgb);
            WriteColor(stream, tableBackgroundShading.FillRgb);
            WriteI16(stream, checked((short)tableBackgroundShading.Pattern));
        }
        if (TableStyleShading is { } tableStyleShading)
        {
            stream.WriteByte(0x87); stream.WriteByte(0xD6); stream.WriteByte(10);
            WriteColor(stream, tableStyleShading.ForegroundRgb);
            WriteColor(stream, tableStyleShading.FillRgb);
            WriteI16(stream, checked((short)tableStyleShading.Pattern));
        }
        if (TableStyleNoWrap is bool tableStyleNoWrap)
        {
            stream.WriteByte(0x7D); stream.WriteByte(0x34);
            stream.WriteByte(tableStyleNoWrap ? (byte)1 : (byte)0);
        }
        if (TableStyleVerticalAlignment is byte tableStyleAlignment)
        {
            if (tableStyleAlignment > 2)
                throw new InvalidDataException("A table-style vertical alignment is invalid.");
            stream.WriteByte(0x7C); stream.WriteByte(0x34);
            stream.WriteByte(tableStyleAlignment);
        }
        if (TableHorizontalBandSize is byte horizontalBandSize)
        {
            if (horizontalBandSize is < 1 or > 3)
                throw new InvalidDataException("A DOC row band size is invalid.");
            stream.WriteByte(0x88); stream.WriteByte(0x34);
            stream.WriteByte(horizontalBandSize);
        }
        if (TableVerticalBandSize is byte verticalBandSize)
        {
            if (verticalBandSize is < 1 or > 3)
                throw new InvalidDataException("A DOC column band size is invalid.");
            stream.WriteByte(0x89); stream.WriteByte(0x34);
            stream.WriteByte(verticalBandSize);
        }
        if (TableAutoFit is bool tableAutoFit)
        {
            stream.WriteByte(0x15); stream.WriteByte(0x36);
            stream.WriteByte(tableAutoFit ? (byte)1 : (byte)0);
        }
        if (TablePreferredWidth is { } preferredWidth)
        {
            if (preferredWidth.Unit is not (1 or 2 or 3) ||
                (preferredWidth.Unit == 1 && preferredWidth.Value != 0) ||
                (preferredWidth.Unit == 2 && preferredWidth.Value > 30000) ||
                (preferredWidth.Unit == 3 && preferredWidth.Value > 31680))
                throw new InvalidDataException("A DOC table has an invalid preferred width.");
            stream.WriteByte(0x14); stream.WriteByte(0xF6);
            stream.WriteByte(preferredWidth.Unit);
            WriteI16(stream, checked((short)preferredWidth.Value));
        }
        foreach (var (code, width) in new[]
        {
            (0x17, TableWidthBefore), (0x18, TableWidthAfter)
        })
        {
            if (width is not { } gap) continue;
            if (gap.Unit is not (1 or 2 or 3) ||
                (gap.Unit == 1 && gap.Value != 0) ||
                (gap.Unit == 2 && gap.Value > 5000) ||
                (gap.Unit == 3 && gap.Value > 31680))
                throw new InvalidDataException("A DOC row gap has an invalid width.");
            stream.WriteByte((byte)code); stream.WriteByte(0xF6);
            stream.WriteByte(gap.Unit);
            WriteI16(stream, checked((short)gap.Value));
        }
        if (TableCellPreferredWidths is { Count: > 0 } preferredCells)
        {
            if (preferredCells.Count > 63)
                throw new InvalidDataException("A DOC table has too many cell widths.");
            for (var i = 0; i < preferredCells.Count; i++)
            {
                if (preferredCells[i] is not { } preferredCell) continue;
                if (preferredCell.Unit is not (1 or 2 or 3) ||
                    (preferredCell.Unit == 1 && preferredCell.Value != 0) ||
                    (preferredCell.Unit == 2 && preferredCell.Value > 5000) ||
                    (preferredCell.Unit == 3 && preferredCell.Value > 31680))
                    throw new InvalidDataException("A DOC cell has an invalid preferred width.");
                stream.WriteByte(0x35); stream.WriteByte(0xD6);
                stream.WriteByte(5); stream.WriteByte(checked((byte)i));
                stream.WriteByte(checked((byte)(i + 1)));
                stream.WriteByte(preferredCell.Unit);
                WriteI16(stream, checked((short)preferredCell.Value));
            }
        }
        if (TableCellNoWraps is { Count: > 0 } noWraps)
        {
            if (noWraps.Count > 63)
                throw new InvalidDataException("A table row has too many cell no-wrap entries.");
            for (var i = 0; i < noWraps.Count; i++)
            {
                if (noWraps[i] is not bool noWrap) continue;
                stream.WriteByte(0x39); stream.WriteByte(0xD6);
                stream.WriteByte(3); stream.WriteByte(checked((byte)i));
                stream.WriteByte(checked((byte)(i + 1)));
                stream.WriteByte(noWrap ? (byte)1 : (byte)0);
            }
        }
        if (TableCellFitTexts is { Count: > 0 } fitTexts)
        {
            if (fitTexts.Count > 63)
                throw new InvalidDataException("A table row has too many cell fit-text entries.");
            for (var i = 0; i < fitTexts.Count; i++)
            {
                if (fitTexts[i] is not bool fitText) continue;
                stream.WriteByte(0x36); stream.WriteByte(0xF6);
                stream.WriteByte(checked((byte)i));
                stream.WriteByte(checked((byte)(i + 1)));
                stream.WriteByte(fitText ? (byte)1 : (byte)0);
            }
        }
        if (TableCellEdges is { Count: > 1 })
        {
            // Word opens generated DOC rows with zero side padding unless
            // defaults are written. DOCX capture supplies the observed
            // 10-twip default for documents without a styles part; other
            // unstyled tables retain this 108-twip fallback.
            var defaultMargins = TableDefaultCellMargins ?? new DocCellMargins();
            WriteCellMargins(stream, 0x34, 0, defaultMargins with
            {
                Left = defaultMargins.Left ?? 108,
                Right = defaultMargins.Right ?? 108
            });
        }
        else if (TableDefaultCellMargins is { } styleMargins)
            WriteCellMargins(stream, 0x34, 0, styleMargins);
        if (TableCellSpacingTwips is ushort cellSpacing)
        {
            if (cellSpacing > 15840)
                throw new InvalidDataException("A DOC cell spacing exceeds 11 inches.");
            stream.WriteByte(0x33); stream.WriteByte(0xD6);
            stream.WriteByte(6); stream.WriteByte(0); stream.WriteByte(1);
            stream.WriteByte(0x0F); stream.WriteByte(3);
            WriteI16(stream, checked((short)cellSpacing));
        }
        if (TableCellMargins is { Count: > 0 } cellMargins)
        {
            if (cellMargins.Count > 63)
                throw new InvalidDataException("A DOC table has too many cell margin entries.");
            for (var i = 0; i < cellMargins.Count; i++)
                if (cellMargins[i] is { } margins)
                    WriteCellMargins(stream, 0x32, i, margins);
        }
        if (TableCantSplit is bool tableCantSplit)
        {
            stream.WriteByte(0x66); stream.WriteByte(0x34);
            stream.WriteByte(tableCantSplit ? (byte)1 : (byte)0);
        }
        if (TableIndentTwips != null || TableRowOriginTwips != null)
        {
            // Word writes the average first-cell side margin as the offset
            // used to position the table's logical left edge.
            var leftMargin = TableDefaultCellMargins?.Left ?? 108;
            var rightMargin = TableDefaultCellMargins?.Right ?? 108;
            if (TableRowOriginTwips is short explicitOrigin)
            {
                stream.WriteByte(0x01); stream.WriteByte(0x96);
                WriteI16(stream, explicitOrigin);
            }
            stream.WriteByte(0x02); stream.WriteByte(0x96);
            WriteI16(stream, checked((short)((leftMargin + rightMargin) / 2)));
        }
        if (TableRowHeightTwips is short tableRowHeight)
        {
            stream.WriteByte(0x07); stream.WriteByte(0x94);
            WriteI16(stream, tableRowHeight);
        }
        if (TableCellEdges is { Count: > 1 } edges)
        {
            if (edges.Count > 64)
                throw new InvalidDataException("The table row has invalid cell edges.");
            for (var i = 1; i < edges.Count; i++)
                if (edges[i] < edges[i - 1])
                    throw new InvalidDataException("The table row has invalid cell edges.");
            var cellCount = edges.Count - 1;
            var merges = TableCellVerticalMerges;
            var horizontalMerges = TableCellHorizontalMerges;
            if (merges != null && merges.Count != cellCount)
                throw new InvalidDataException("The table row has inconsistent cell merges.");
            if (horizontalMerges != null && horizontalMerges.Count != cellCount)
                throw new InvalidDataException("The table row has inconsistent horizontal cell merges.");
            var hasMerges = merges != null || horizontalMerges != null;
            var hasCellDefinitions = hasMerges ||
                TableCellBorders?.Any(x => x != null) == true ||
                TableCellTextFlows?.Any(x => x != null) == true ||
                TableCellHideMarks?.Any(x => x == true) == true ||
                TableCellFitTexts?.Any(x => x == true) == true ||
                TableCellNoWraps?.Any(x => x == true) == true;
            stream.WriteByte(0x08); stream.WriteByte(0xD6);
            WriteI16(stream, checked((short)(2 + edges.Count * 2 +
                (hasCellDefinitions ? cellCount * 20 : 0))));
            stream.WriteByte(checked((byte)(edges.Count - 1)));
            foreach (var edge in edges) WriteI16(stream, edge);
            if (hasCellDefinitions)
                for (var i = 0; i < cellCount; i++)
                {
                    var merge = merges?[i];
                    var horizontal = horizontalMerges?[i];
                    if (merge is not (null or 1 or 3))
                        throw new InvalidDataException("A DOC table cell has invalid vertical merge flags.");
                    if (horizontal is not (null or 1 or 2 or 3))
                        throw new InvalidDataException("A DOC table cell has invalid horizontal merge flags.");
                    WriteI16(stream, checked((short)((horizontal ?? 0) |
                        ((merge ?? 0) << 5) |
                        ((TableCellTextFlows != null && i < TableCellTextFlows.Count
                            ? TableCellTextFlows[i] ?? 0 : 0) << 2) |
                        (TableCellFitTexts != null && i < TableCellFitTexts.Count &&
                            TableCellFitTexts[i] == true ? 0x1000 : 0) |
                        (TableCellNoWraps != null && i < TableCellNoWraps.Count &&
                            TableCellNoWraps[i] == true ? 0x2000 : 0) |
                        (TableCellHideMarks != null && i < TableCellHideMarks.Count &&
                            TableCellHideMarks[i] == true ? 0x4000 : 0))));
                    stream.WriteByte(0); stream.WriteByte(0);
                    var borders = TableCellBorders != null && i < TableCellBorders.Count
                        ? TableCellBorders[i] : null;
                    WriteCellDefinitionBorder(stream, borders?.Top);
                    WriteCellDefinitionBorder(stream, borders?.Left);
                    WriteCellDefinitionBorder(stream, borders?.Bottom);
                    WriteCellDefinitionBorder(stream, borders?.Right);
                }
        }
        // TDefTable establishes the row's cell origins. Word writes the
        // preferred table indent after that definition so it remains effective.
        if (TableIndentTwips is short tableIndent)
        {
            if (tableIndent is < -31560 or > 31680)
                throw new InvalidDataException("A DOC table indent exceeds the DOC limit.");
            stream.WriteByte(0x61); stream.WriteByte(0xF6);
            stream.WriteByte(3);
            WriteI16(stream, tableIndent);
        }
        if (TableJustification is byte tableJustification)
        {
            if (tableJustification > 2)
                throw new InvalidDataException("A DOC table has invalid justification.");
            stream.WriteByte(0x00); stream.WriteByte(0x54);
            WriteI16(stream, tableJustification);
            stream.WriteByte(0x8A); stream.WriteByte(0x54);
            WriteI16(stream, tableJustification);
        }
        if (TableCellShadings is { Count: > 0 } cellShadings)
        {
            if (cellShadings.Count > 63)
                throw new InvalidDataException("A table row has too many cell shading entries.");
            static byte? LegacyCellColor(uint? value) => value switch
            {
                null or 0xFF000000u => 0,
                0x000000u => 1, 0xFF0000u => 2, 0xFFFF00u => 3,
                0x00FF00u => 4, 0xFF00FFu => 5, 0x0000FFu => 6,
                0x00FFFFu => 7, 0xFFFFFFu => 8, 0x800000u => 9,
                0x808000u => 10, 0x008000u => 11,
                0x800080u => 12, 0x008080u => 14,
                0x808080u => 15, 0xC0C0C0u => 16,
                _ => null
            };
            var legacyShadings = new ushort[cellShadings.Count];
            var hasLegacyShading = false;
            for (var i = 0; i < cellShadings.Count; i++)
            {
                var shading = cellShadings[i];
                if (shading == null || shading.Pattern > 63 ||
                    LegacyCellColor(shading.ForegroundRgb) is not byte foreground ||
                    LegacyCellColor(shading.FillRgb) is not byte background)
                {
                    legacyShadings[i] = TableStyleIndex != null ? (ushort)0xFFFF : (ushort)0;
                    continue;
                }
                legacyShadings[i] = (ushort)(foreground | (background << 5) |
                    (shading.Pattern << 10));
                hasLegacyShading = true;
            }
            if (hasLegacyShading)
            {
                stream.WriteByte(0x09); stream.WriteByte(0xD6);
                stream.WriteByte(checked((byte)(legacyShadings.Length * 2)));
                foreach (var legacy in legacyShadings)
                    WriteI16(stream, unchecked((short)legacy));
            }
            for (var first = 0; first < cellShadings.Count; first += 22)
            {
                var count = Math.Min(22, cellShadings.Count - first);
                stream.WriteByte((byte)(0x70 + first / 22)); stream.WriteByte(0xD6);
                stream.WriteByte(checked((byte)(count * 10)));
                for (var i = 0; i < count; i++)
                {
                    var shading = cellShadings[first + i];
                    if (shading?.Pattern == ushort.MaxValue)
                    {
                        for (var colorByte = 0; colorByte < 4; colorByte++)
                            stream.WriteByte(0xFF);
                        WriteColor(stream, null);
                        WriteI16(stream, 0);
                    }
                    else
                    {
                        WriteColor(stream, shading?.ForegroundRgb);
                        WriteColor(stream, shading?.FillRgb);
                        WriteI16(stream, checked((short)(shading?.Pattern ?? 0)));
                    }
                }
            }
        }
        if (TableCellVerticalAlignments is { Count: > 0 } alignments)
        {
            if (alignments.Count > 63)
                throw new InvalidDataException("A table row has too many cell alignment entries.");
            for (var i = 0; i < alignments.Count; i++)
            {
                if (alignments[i] is not byte alignment) continue;
                if (alignment > 2)
                    throw new InvalidDataException("A table cell vertical alignment is invalid.");
                stream.WriteByte(0x2C); stream.WriteByte(0xD6);
                stream.WriteByte(3); stream.WriteByte(checked((byte)i));
                stream.WriteByte(checked((byte)(i + 1))); stream.WriteByte(alignment);
            }
        }
        if (TableCellTextFlows is { Count: > 0 } textFlows)
        {
            if (textFlows.Count > 63)
                throw new InvalidDataException("A table row has too many cell text flow entries.");
            for (var i = 0; i < textFlows.Count; i++)
            {
                if (textFlows[i] is not ushort textFlow) continue;
                if (textFlow is not (0 or 1 or 3 or 4 or 5))
                    throw new InvalidDataException("A table cell text flow is invalid.");
                stream.WriteByte(0x29); stream.WriteByte(0x76);
                stream.WriteByte(checked((byte)i));
                stream.WriteByte(checked((byte)(i + 1)));
                WriteI16(stream, checked((short)textFlow));
            }
        }
        if (TableCellHideMarks is { Count: > 0 } hideMarks)
        {
            if (hideMarks.Count > 63)
                throw new InvalidDataException("A table row has too many cell hide-mark entries.");
            for (var i = 0; i < hideMarks.Count; i++)
            {
                if (hideMarks[i] is not bool hideMark) continue;
                stream.WriteByte(0x42); stream.WriteByte(0xD6);
                stream.WriteByte(3); stream.WriteByte(checked((byte)i));
                stream.WriteByte(checked((byte)(i + 1)));
                stream.WriteByte(hideMark ? (byte)1 : (byte)0);
            }
        }
        if (TableBorders is { } tableBorders)
        {
            stream.WriteByte(0x05); stream.WriteByte(0xD6); stream.WriteByte(24);
            WriteLegacyTableBorder(stream, tableBorders.Top);
            WriteLegacyTableBorder(stream, tableBorders.Left);
            WriteLegacyTableBorder(stream, tableBorders.Bottom);
            WriteLegacyTableBorder(stream, tableBorders.Right);
            WriteLegacyTableBorder(stream, tableBorders.InsideHorizontal);
            WriteLegacyTableBorder(stream, tableBorders.InsideVertical);
            stream.WriteByte(0x13); stream.WriteByte(0xD6); stream.WriteByte(48);
            WriteTableBorder(stream, tableBorders.Top);
            WriteTableBorder(stream, tableBorders.Left);
            WriteTableBorder(stream, tableBorders.Bottom);
            WriteTableBorder(stream, tableBorders.Right);
            WriteTableBorder(stream, tableBorders.InsideHorizontal);
            WriteTableBorder(stream, tableBorders.InsideVertical);
        }
        if (TableCellBorders is { Count: > 0 } cellBorders)
        {
            if (cellBorders.Count > 63)
                throw new InvalidDataException("A table row has too many cell border entries.");
            // A border operand addresses a range of cells. Coalesce equal
            // adjacent edges rather than repeating a modifier per cell.
            foreach (var (side, select) in new (byte Side,
                Func<DocCellBorders, DocParagraphBorder?> Select)[]
            {
                (1, x => x.Top), (2, x => x.Left),
                (4, x => x.Bottom), (8, x => x.Right),
                (0x10, x => x.TopLeftToBottomRight),
                (0x20, x => x.TopRightToBottomLeft)
            })
            {
                for (var i = 0; i < cellBorders.Count;)
                {
                    var border = cellBorders[i] is { } cell ? select(cell) : null;
                    if (border == null) { i++; continue; }
                    var end = i + 1;
                    while (end < cellBorders.Count &&
                        cellBorders[end] is { } next && select(next) == border)
                        end++;
                    WriteCellBorder(stream, i, end, side, border);
                    i = end;
                }
            }
        }
        if (LineValue is short line)
        {
            stream.WriteByte(0x12); stream.WriteByte(0x64);
            stream.WriteByte((byte)line); stream.WriteByte((byte)(line >> 8));
            stream.WriteByte(LineIsMultiple == true ? (byte)1 : (byte)0);
            stream.WriteByte(0);
        }
        if (BeforeTwips is ushort before)
        {
            stream.WriteByte(0x13); stream.WriteByte(0xA4);
            stream.WriteByte((byte)before); stream.WriteByte((byte)(before >> 8));
        }
        if (AfterTwips is ushort after)
        {
            stream.WriteByte(0x14); stream.WriteByte(0xA4);
            stream.WriteByte((byte)after); stream.WriteByte((byte)(after >> 8));
        }
        if (BeforeLines is short beforeLines)
        {
            stream.WriteByte(0x58); stream.WriteByte(0x44);
            WriteI16(stream, beforeLines);
        }
        if (AfterLines is short afterLines)
        {
            stream.WriteByte(0x59); stream.WriteByte(0x44);
            WriteI16(stream, afterLines);
        }
        if (RightChars is short rightChars)
        {
            stream.WriteByte(0x55); stream.WriteByte(0x44);
            WriteI16(stream, rightChars);
        }
        if (LeftChars is short leftChars)
        {
            stream.WriteByte(0x56); stream.WriteByte(0x44);
            WriteI16(stream, leftChars);
        }
        if (FirstLineChars is short firstLineChars)
        {
            stream.WriteByte(0x57); stream.WriteByte(0x44);
            WriteI16(stream, firstLineChars);
        }
        if (RightTwips is short right) WriteSigned(stream, 0x5D, right);
        if (LeftTwips is short left) WriteSigned(stream, 0x5E, left);
        if (FirstLineTwips is short firstLine) WriteSigned(stream, 0x60, firstLine);
        if ((ClearedTabPositions?.Count ?? 0) > 0 || (TabStops?.Count ?? 0) > 0)
            WriteTabs(stream, forStyle);
        WriteBorder(stream, 0x4E, TopBorder);
        WriteBorder(stream, 0x4F, LeftBorder);
        WriteBorder(stream, 0x50, BottomBorder);
        WriteBorder(stream, 0x51, RightBorder);
        WriteBorder(stream, 0x52, BetweenBorder);
        if (FillRgb != null || ShadingForegroundRgb != null || ShadingPattern != null)
        {
            static byte? LegacyShadingIndex(uint? value) => value switch
            {
                null or 0xFF000000u => 0,
                0x000000u => 1, 0xFF0000u => 2, 0xFFFF00u => 3,
                0x00FF00u => 4, 0xFF00FFu => 5, 0x0000FFu => 6,
                0x00FFFFu => 7, 0xFFFFFFu => 8, 0x800000u => 9,
                0x808000u => 10, 0x008000u => 11,
                0x800080u => 12, 0x008080u => 14,
                0x808080u => 15, 0xC0C0C0u => 16,
                _ => null
            };
            var legacyForeground = LegacyShadingIndex(ShadingForegroundRgb);
            var legacyBackground = LegacyShadingIndex(FillRgb);
            var pattern = ShadingPattern ?? 0;
            var nil = pattern == ushort.MaxValue;
            if (!nil && legacyForeground is byte foreground &&
                legacyBackground is byte background && pattern <= 63)
            {
                var shd80 = (ushort)(foreground | (background << 5) |
                    (pattern << 10));
                stream.WriteByte(0x2D); stream.WriteByte(0x44);
                stream.WriteByte((byte)shd80); stream.WriteByte((byte)(shd80 >> 8));
            }
            stream.WriteByte(0x4D); stream.WriteByte(0xC6); stream.WriteByte(10);
            WriteColor(stream, nil ? null : ShadingForegroundRgb);
            WriteColor(stream, nil ? null : FillRgb);
            stream.WriteByte(nil ? (byte)0 : (byte)pattern);
            stream.WriteByte(nil ? (byte)0 : (byte)(pattern >> 8));
        }
        return stream.ToArray();
    }

    private static DocParagraphBorder? ParseOptionalCellBorder80(ReadOnlySpan<byte> border) =>
        border[0] == 0 && border[1] == 0 && border[2] == 0 && border[3] == 0
            ? null
            : border[0] == 0xFF && border[1] == 0xFF &&
                border[2] == 0xFF && border[3] == 0xFF
                ? new DocParagraphBorder(0, 0, 0, null)
                : DocParagraphBorder.Parse80(border);

    private static void ParseTabs(ReadOnlySpan<byte> operand, List<short> cleared,
        List<DocTabClearRange> clearRanges, List<DocTabStop> added,
        bool withTolerance)
    {
        if (operand.Length < 3 ||
            (operand[0] == 255 ? !withTolerance : operand[0] + 1 != operand.Length))
            throw new InvalidDataException("A custom tab operand has an invalid length.");
        var deletedCount = operand[1];
        var deletionBytes = deletedCount * (withTolerance ? 4 : 2);
        if (deletedCount > 64 || 2 + deletionBytes >= operand.Length)
            throw new InvalidDataException("A custom tab deletion list is invalid.");
        for (var i = 0; i < deletedCount; i++)
        {
            var position = BinaryPrimitives.ReadInt16LittleEndian(operand.Slice(2 + i * 2));
            cleared.Add(position);
            if (withTolerance)
            {
                var tolerance = BinaryPrimitives.ReadUInt16LittleEndian(
                    operand.Slice(2 + deletedCount * 2 + i * 2));
                clearRanges.Add(new DocTabClearRange(position,
                    Math.Max((ushort)25, tolerance)));
            }
        }
        var addedCountOffset = 2 + deletionBytes;
        var addedCount = operand[addedCountOffset];
        if (addedCount > 64 || addedCountOffset + 1 + addedCount * 3 != operand.Length)
            throw new InvalidDataException("A custom tab addition list is invalid.");
        var positionsOffset = addedCountOffset + 1;
        var descriptorsOffset = positionsOffset + addedCount * 2;
        for (var i = 0; i < addedCount; i++)
        {
            var position = BinaryPrimitives.ReadInt16LittleEndian(
                operand.Slice(positionsOffset + i * 2));
            var descriptor = operand[descriptorsOffset + i];
            added.Add(new DocTabStop(position, (byte)(descriptor & 7),
                (byte)((descriptor >> 3) & 7)));
        }
    }

    internal static int GetExtendedTabOperandLength(ReadOnlySpan<byte> operand)
    {
        if (operand.Length < 3)
            throw new InvalidDataException("A custom tab operand is truncated.");
        var deletedCount = operand[1];
        if (deletedCount > 64 || operand.Length <= 2 + deletedCount * 4)
            throw new InvalidDataException("A custom tab deletion list is invalid.");
        var addedCount = operand[2 + deletedCount * 4];
        if (addedCount > 64)
            throw new InvalidDataException("A custom tab addition list is invalid.");
        return 3 + deletedCount * 4 + addedCount * 3;
    }

    private void WriteTabs(Stream stream, bool forStyle)
    {
        var cleared = (ClearedTabPositions ?? Array.Empty<short>()).OrderBy(x => x).ToArray();
        var added = (TabStops ?? Array.Empty<DocTabStop>()).OrderBy(x => x.PositionTwips).ToArray();
        var length = 2 + cleared.Length * (forStyle ? 4 : 2) + added.Length * 3;
        if (cleared.Length > 64 || added.Length > 64 || length > 255 ||
            cleared.Distinct().Count() != cleared.Length ||
            added.Select(x => x.PositionTwips).Distinct().Count() != added.Length)
            throw new InvalidDataException("The custom tab list is too large or has duplicate positions.");
        stream.WriteByte(forStyle ? (byte)0x15 : (byte)0x0D);
        stream.WriteByte(0xC6);
        stream.WriteByte(checked((byte)length));
        stream.WriteByte(checked((byte)cleared.Length));
        foreach (var position in cleared) WriteI16(stream, position);
        if (forStyle)
            foreach (var _ in cleared) WriteI16(stream, 25);
        stream.WriteByte(checked((byte)added.Length));
        foreach (var tab in added) WriteI16(stream, tab.PositionTwips);
        foreach (var tab in added)
            stream.WriteByte(checked((byte)((tab.Alignment & 7) | ((tab.Leader & 7) << 3))));
    }

    private static void WriteI16(Stream stream, short value)
    {
        stream.WriteByte((byte)value); stream.WriteByte((byte)(value >> 8));
    }

    private static void WriteU32(Stream stream, uint value)
    {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)(value >> 16));
        stream.WriteByte((byte)(value >> 24));
    }

    private static void WriteCellBorder(Stream stream, int index, int end, byte side,
        DocParagraphBorder? border)
    {
        if (border == null) return;
        // The modern border operand retains the complete color and line style.
        // Emitting the legacy operand as well can overflow a row PAPX and cause
        // Word to discard that row's table properties.
        stream.WriteByte(0x2F); stream.WriteByte(0xD6);
        stream.WriteByte(11); stream.WriteByte(checked((byte)index));
        stream.WriteByte(checked((byte)end)); stream.WriteByte(side);
        var bytes = border.EncodeRaw();
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteCellDefinitionBorder(Stream stream,
        DocParagraphBorder? border)
    {
        var bytes = border?.Type is 0 or 26 or 27
            ? new byte[] { 0xFF, 0xFF, 0xFF, 0xFF }
            : border?.Encode80() ?? new byte[4];
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteCellMargins(Stream stream, byte opcode, int index,
        DocCellMargins margins)
    {
        WriteCellMargin(stream, opcode, index, 1, margins.Top);
        WriteCellMargin(stream, opcode, index, 2, margins.Left);
        WriteCellMargin(stream, opcode, index, 4, margins.Bottom);
        WriteCellMargin(stream, opcode, index, 8, margins.Right);
    }

    private static void WriteCellMargin(Stream stream, byte opcode, int index,
        byte side, ushort? margin)
    {
        if (margin == null) return;
        if (margin > 31680)
            throw new InvalidDataException("A DOC cell margin exceeds 22 inches.");
        stream.WriteByte(opcode); stream.WriteByte(0xD6);
        stream.WriteByte(6); stream.WriteByte(checked((byte)index));
        stream.WriteByte(checked((byte)(index + 1)));
        stream.WriteByte(side); stream.WriteByte(3);
        WriteI16(stream, checked((short)margin.Value));
    }

    private static void WriteTableBorder(Stream stream, DocParagraphBorder? border)
    {
        var bytes = border?.EncodeRaw() ?? new byte[8];
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteLegacyTableBorder(Stream stream, DocParagraphBorder? border)
    {
        var bytes = border?.Encode80() ?? new byte[] { 0xFF, 0xFF, 0xFF, 0xFF };
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteColor(Stream stream, uint? color)
    {
        var value = color ?? 0;
        stream.WriteByte((byte)value); stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)(value >> 16)); stream.WriteByte(color == null ? (byte)0xFF : (byte)0);
    }

    private static void WriteBorder(Stream stream, byte opcode, DocParagraphBorder? border)
    {
        if (border == null) return;
        if (border.Encode80() is { } legacy)
        {
            stream.WriteByte((byte)(opcode - 0x2A)); stream.WriteByte(0x64);
            stream.Write(legacy, 0, legacy.Length);
        }
        stream.WriteByte(opcode); stream.WriteByte(0xC6);
        var bytes = border.Encode();
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteSigned(Stream stream, byte opcode, short value)
    {
        stream.WriteByte(opcode); stream.WriteByte(0x84);
        stream.WriteByte((byte)value); stream.WriteByte((byte)(value >> 8));
    }

    private static void WriteBool(Stream stream, byte opcode, bool value)
    {
        stream.WriteByte(opcode); stream.WriteByte(0x24); stream.WriteByte(value ? (byte)1 : (byte)0);
    }
}
