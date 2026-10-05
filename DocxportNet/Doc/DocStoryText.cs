using System.Globalization;
using System.Text;

namespace DocxportNet.Doc;

/// <summary>A logical text story with source character positions.</summary>
public sealed record DocStoryText(string Name, uint CpStart, uint CpEnd,
    IReadOnlyList<DocStoryParagraph> Paragraphs);

/// <summary>A logical story with its overlapping source formatting ranges.</summary>
public sealed record DocIndexedStoryText(DocStoryText Text,
    IReadOnlyList<DocCharacterFormattingRange> CharacterFormatting,
    IReadOnlyList<DocParagraphStyleRange> ParagraphStyles)
{
    public IReadOnlyList<DocStoryParagraphRange> ParagraphRanges { get; init; } = [];
    public IReadOnlyList<DocStoryCharacterRun> Runs { get; init; } = [];
    public IReadOnlyList<DocStoryParagraphStyleRun> StoryParagraphStyles { get; init; } = [];
    public IReadOnlyList<DocStoryTableRow> TableRows { get; init; } = [];
    public IReadOnlyList<DocStoryListParagraph> ListParagraphs { get; init; } = [];
    public IReadOnlyList<DocStoryFieldMark> FieldMarks { get; init; } = [];
    public IReadOnlyList<DocStoryBookmark> Bookmarks { get; init; } = [];
    public IReadOnlyDictionary<int, DocStoryInlinePicture> InlinePictures { get; init; } =
        new Dictionary<int, DocStoryInlinePicture>();
    public IReadOnlyDictionary<int, DocStoryFloatingPicture[]> FloatingPictures { get; init; } =
        new Dictionary<int, DocStoryFloatingPicture[]>();
}

/// <summary>A paragraph's story-relative bounds and terminating control.</summary>
public sealed record DocStoryParagraphRange(int Start, int End, DocParagraphEnd EndKind,
    uint? SourceCpStart = null, uint? SourceCpEnd = null,
    long? SourceTextOffsetStart = null, long? SourceTextOffsetEnd = null)
{
    internal static IReadOnlyList<DocStoryParagraphRange> Capture(string text,
        IReadOnlyList<DocPlainTextParagraphStyleRun> styles,
        IReadOnlyCollection<int>? sectionEnds = null)
    {
        var ranges = new List<DocStoryParagraphRange>();
        var start = 0;
        for (var cp = 0; cp < text.Length; cp++)
        {
            var character = text[cp];
            if (character is not ('\r' or '\u0007' or '\f')) continue;
            var formatting = styles.FirstOrDefault(x => x.Start <= cp && cp < x.End)?
                .Formatting;
            if (character == '\u0007' && formatting?.InTable != true ||
                character == '\f' && sectionEnds?.Contains(cp + 1) != true)
                continue;
            var end = character == '\f' ? DocParagraphEnd.SectionMark :
                character == '\u0007'
                    ? formatting?.TableTerminator == true ? DocParagraphEnd.RowMark
                        : DocParagraphEnd.CellMark
                    : formatting?.InnerTableRow == true ? DocParagraphEnd.RowMark
                        : formatting?.InnerTableCell == true ? DocParagraphEnd.CellMark
                        : DocParagraphEnd.ParagraphMark;
            ranges.Add(new DocStoryParagraphRange(start, cp + 1, end));
            start = cp + 1;
        }
        if (start < text.Length || ranges.Count == 0)
            ranges.Add(new DocStoryParagraphRange(start, text.Length,
                DocParagraphEnd.None));
        return ranges;
    }

    internal static void Validate(string text,
        IReadOnlyList<DocStoryParagraphRange> paragraphs)
    {
        var start = 0;
        foreach (var paragraph in paragraphs)
        {
            var terminal = paragraph.End > start && paragraph.End <= text.Length
                ? text[paragraph.End - 1] : '\0';
            var validEnd = paragraph.EndKind switch
            {
                DocParagraphEnd.None => paragraph.End == text.Length,
                DocParagraphEnd.ParagraphMark => terminal == '\r',
                DocParagraphEnd.SectionMark => terminal == '\f',
                DocParagraphEnd.CellMark or DocParagraphEnd.RowMark =>
                    terminal is '\r' or '\u0007',
                _ => false
            };
            if (paragraph.Start != start || paragraph.End < start ||
                paragraph.End > text.Length || !validEnd)
                throw new InvalidDataException("A captured paragraph exceeds its story or has an invalid end.");
            start = paragraph.End;
        }
        if (paragraphs.Count == 0 || start != text.Length)
            throw new InvalidDataException("Captured paragraphs do not cover their story.");
    }
}

/// <summary>A character formatting range in story coordinates.</summary>
public sealed record DocStoryCharacterRun(int Start, int End,
    DocCharacterFormatting Formatting, uint? SourceCpStart = null,
    uint? SourceCpEnd = null, long? SourceTextOffsetStart = null,
    long? SourceTextOffsetEnd = null)
{
    internal DocPlainTextFormatRun ToWriter(int offset) => new(
        checked(offset + Start), checked(offset + End), Formatting);
}

/// <summary>A paragraph style and direct formatting range in story coordinates.</summary>
public sealed record DocStoryParagraphStyleRun(int Start, int End, int StyleIndex,
    DocParagraphFormatting? Formatting = null, uint? SourceCpStart = null,
    uint? SourceCpEnd = null, long? SourceTextOffsetStart = null,
    long? SourceTextOffsetEnd = null)
{
    internal DocPlainTextParagraphStyleRun ToWriter(int offset) => new(
        checked(offset + Start), checked(offset + End), StyleIndex, Formatting);
}

/// <summary>A table row's paragraph formatting in story coordinates.</summary>
public sealed record DocStoryTableRow(int Start, int End, int StyleIndex,
    DocParagraphFormatting Formatting, uint? SourceCpStart = null,
    uint? SourceCpEnd = null, long? SourceTextOffset = null)
{
    public static bool IsRow(DocParagraphFormatting? formatting) =>
        formatting?.InTable == true &&
        (formatting.TableTerminator == true || formatting.InnerTableRow == true) &&
        formatting.TableCellEdges is { Count: > 1 };
}

/// <summary>A directly numbered paragraph range in story coordinates.</summary>
public sealed record DocStoryListParagraph(int Start, int End, short OverrideIndex,
    byte Level, uint? SourceCpStart = null, uint? SourceCpEnd = null,
    long? SourceTextOffsetStart = null, long? SourceTextOffsetEnd = null)
{
    internal static IReadOnlyList<DocStoryListParagraph> Capture(
        IEnumerable<DocStoryParagraphStyleRun> paragraphs,
        IEnumerable<DocStoryTableRow> rows) =>
        paragraphs.Where(x => x.Formatting?.ListOverrideIndex > 0)
            .Select(x => new DocStoryListParagraph(x.Start, x.End,
                x.Formatting!.ListOverrideIndex!.Value, x.Formatting.ListLevel ?? 0))
            .Concat(rows.Where(x => x.Formatting.ListOverrideIndex > 0)
                .Select(x => new DocStoryListParagraph(x.Start, x.End,
                    x.Formatting.ListOverrideIndex!.Value,
                    x.Formatting.ListLevel ?? 0)))
            .OrderBy(x => x.Start).ThenBy(x => x.End).ToArray();

    internal static void Validate(IReadOnlyList<DocStoryListParagraph> lists,
        IEnumerable<DocStoryParagraphStyleRun> paragraphs,
        IEnumerable<DocStoryTableRow> rows)
    {
        var expected = Capture(paragraphs, rows);
        if (lists.Count != expected.Count || lists.Where((x, i) =>
            x.Start != expected[i].Start || x.End != expected[i].End ||
            x.OverrideIndex != expected[i].OverrideIndex ||
            x.Level != expected[i].Level).Any())
            throw new InvalidDataException(
                "A captured list paragraph differs from its paragraph formatting.");
    }
}

/// <summary>A field control at a story-relative position with optional source DOC CP.</summary>
public sealed record DocStoryFieldMark(int Cp, byte Kind, byte FieldType,
    uint? SourceCp = null, long? SourceTextOffset = null);

/// <summary>A DOC field PLC marker in global CP space; type is recorded on begins.</summary>
public sealed record DocIndexedFieldMark(uint Cp, byte Kind, byte FieldType);

/// <summary>A bookmark in story coordinates with optional source DOC positions.</summary>
public sealed record DocStoryBookmark(string Name, int Start, int End,
    int? Id = null, uint? SourceCpStart = null, uint? SourceCpEnd = null,
    long? SourceTextOffsetStart = null, long? SourceTextOffsetEnd = null);

/// <summary>An inline picture at a story-relative position with optional source DOC coordinates.</summary>
public sealed record DocStoryInlinePicture(int Cp, int? DataOffset = null,
    uint? SourceCp = null, long? SourceTextOffset = null)
{
    internal DocInlinePicture? Payload { get; init; }
}

/// <summary>A floating image and its layout in story coordinates.</summary>
public sealed record DocStoryFloatingPicture(int Cp, byte[] Bytes, string ContentType,
    long WidthEmu, long HeightEmu, int LeftTwips, int TopTwips,
    byte WrapCode = 2, bool BehindText = false, byte WrapSide = 0,
    byte HorizontalOrigin = 0, byte VerticalOrigin = 0,
    byte HorizontalAlignment = 0, byte VerticalAlignment = 0,
    int DistanceTopEmu = 0, int DistanceBottomEmu = 0,
    int DistanceLeftEmu = 114300, int DistanceRightEmu = 114300)
{
    public DocInlinePictureCrop? Crop { get; init; }
    public bool FlipHorizontal { get; init; }
    public bool FlipVertical { get; init; }
    public double RotationDegrees { get; init; }
    public uint? ShapeId { get; init; }
    public uint? SourceCp { get; init; }
    public long? SourceTextOffset { get; init; }

    internal DocPlainTextFloatingPicture ToWriter(int cp) => new(cp,
        new DocInlinePicture(Bytes, ContentType, WidthEmu, HeightEmu, Crop,
            FlipHorizontal, FlipVertical, RotationDegrees),
        LeftTwips, TopTwips, WrapCode, BehindText, WrapSide,
        HorizontalOrigin, VerticalOrigin, HorizontalAlignment, VerticalAlignment,
        DistanceTopEmu, DistanceBottomEmu, DistanceLeftEmu, DistanceRightEmu);
}

/// <summary>A paragraph's end position includes its terminating DOC control, if any.</summary>
public sealed record DocStoryParagraph(uint CpStart, uint CpEnd,
    DocParagraphEnd End, IReadOnlyList<DocStoryAtom> Atoms);

public enum DocParagraphEnd { None, ParagraphMark, SectionMark, CellMark, RowMark }

public enum DocStoryAtomKind
{
    Text, Tab, LineBreak, PageBreak, ColumnBreak, CellMark, NoBreakHyphen, SoftHyphen,
    FieldBegin, FieldSeparator, FieldEnd, FieldData, InlinePicture,
    FloatingShapeAnchor, LegacyDateBlock, LegacyPageNumberBlock, UnsupportedControl
}

/// <summary>Text or a single control, located in the original DOC's CP space.</summary>
public sealed record DocStoryAtom(uint CpStart, uint CpEnd, DocStoryAtomKind Kind, string Text);

/// <summary>Assembles piece-table text into logical story paragraphs before projection.</summary>
public static class DocStoryTextReader
{
    public static DocStoryText Read(DocTextIndex index, string partName)
    {
        if (index == null) throw new ArgumentNullException(nameof(index));
        if (partName == null) throw new ArgumentNullException(nameof(partName));
        var part = index.Parts.FirstOrDefault(x => x.Name == partName)
            ?? throw new ArgumentException($"The DOC has no '{partName}' story.", nameof(partName));
        return ReadRange(index, partName, part.CpStart, part.CpEnd);
    }

    /// <summary>Reads a substory in global CP coordinates, excluding any caller-owned guard mark.</summary>
    public static DocStoryText ReadRange(DocTextIndex index, string partName, uint cpStart, uint cpEnd)
    {
        if (index == null) throw new ArgumentNullException(nameof(index));
        if (partName == null) throw new ArgumentNullException(nameof(partName));
        var part = index.Parts.FirstOrDefault(x => x.Name == partName)
            ?? throw new ArgumentException($"The DOC has no '{partName}' story.", nameof(partName));
        if (cpStart < part.CpStart || cpEnd > part.CpEnd || cpStart > cpEnd)
            throw new ArgumentOutOfRangeException(nameof(cpStart), "The story range lies outside its document part.");
        var spans = index.GetPartSpans(partName)
            .Where(x => x.CpStart < cpEnd && x.CpEnd > cpStart)
            .Select(x => new DocIndexedTextSpan(x.Piece,
                Math.Max(x.CpStart, cpStart), Math.Min(x.CpEnd, cpEnd)))
            .ToArray();
        var expected = cpStart;
        var paragraphs = new List<DocStoryParagraph>();
        var atoms = new List<DocStoryAtom>();
        var text = new StringBuilder();
        var paragraphStart = cpStart;
        var textStart = cpStart;
        uint? highSurrogateCp = null;
        char highSurrogate = default;
        var sectionEnds = new Lazy<HashSet<uint>>(() => partName == "Main"
            ? new HashSet<uint>(index.Sections.Take(Math.Max(0, index.Sections.Count - 1))
                .Select(x => uint.Parse(x.Attributes["cpEnd"], CultureInfo.InvariantCulture)))
            : []);
        var fieldMarks = new Lazy<IReadOnlyDictionary<uint, DocIndexedFieldMark>>(() =>
            index.GetFieldRecords(partName).ToDictionary(x => x.Cp));
        var tableParagraphs = new Lazy<IReadOnlyList<DocParagraphStyleRange>>(
            () => index.ParagraphStyles);

        DocParagraphFormatting? FormattingAt(uint cp)
        {
            foreach (var range in tableParagraphs.Value)
                if (range.CpStart <= cp && cp < range.CpEnd) return range.Formatting;
            return null;
        }

        void FlushText(uint cp)
        {
            if (text.Length == 0) return;
            atoms.Add(new DocStoryAtom(textStart, cp, DocStoryAtomKind.Text, text.ToString()));
            text.Clear();
        }

        void AddText(char value, uint cp)
        {
            if (text.Length == 0) textStart = cp;
            text.Append(value);
        }

        void AddControl(DocStoryAtomKind kind, char value, uint cp)
        {
            FlushText(cp);
            atoms.Add(new DocStoryAtom(cp, cp + 1, kind, value.ToString()));
        }

        void EndParagraph(uint cp, DocParagraphEnd end)
        {
            FlushText(cp);
            paragraphs.Add(new DocStoryParagraph(paragraphStart, cp + 1, end, atoms.ToArray()));
            atoms.Clear();
            paragraphStart = cp + 1;
        }

        foreach (var span in spans)
        {
            if (span.CpStart != expected)
                throw new InvalidDataException($"The '{partName}' text index has a gap or overlap.");
            var value = span.Text;
            for (var i = 0; i < value.Length; i++)
            {
                var cp = span.CpStart + (uint)i;
                var character = value[i];
                if (highSurrogateCp is uint highCp)
                {
                    highSurrogateCp = null;
                    if (char.IsLowSurrogate(character))
                    {
                        AddText(highSurrogate, highCp);
                        AddText(character, cp);
                        continue;
                    }
                    AddControl(DocStoryAtomKind.UnsupportedControl, highSurrogate, highCp);
                }
                if (char.IsHighSurrogate(character))
                {
                    highSurrogate = character;
                    highSurrogateCp = cp;
                    continue;
                }
                if (char.IsLowSurrogate(character))
                {
                    AddControl(DocStoryAtomKind.UnsupportedControl, character, cp);
                    continue;
                }
                switch (character)
                {
                    case '\r':
                        var paragraphFormatting = FormattingAt(cp);
                        EndParagraph(cp, paragraphFormatting?.InnerTableRow == true
                            ? DocParagraphEnd.RowMark :
                            paragraphFormatting?.InnerTableCell == true
                                ? DocParagraphEnd.CellMark :
                                DocParagraphEnd.ParagraphMark);
                        break;
                    case '\f' when sectionEnds.Value.Contains(cp + 1):
                        EndParagraph(cp, DocParagraphEnd.SectionMark); break;
                    case '\f': AddControl(DocStoryAtomKind.PageBreak, character, cp); break;
                    case '\u000E': AddControl(DocStoryAtomKind.ColumnBreak, character, cp); break;
                    case '\t': AddControl(DocStoryAtomKind.Tab, character, cp); break;
                    case '\v': AddControl(DocStoryAtomKind.LineBreak, character, cp); break;
                    case '\u0007' when FormattingAt(cp)?.InTable == true:
                        EndParagraph(cp, FormattingAt(cp)?.TableTerminator == true
                            ? DocParagraphEnd.RowMark : DocParagraphEnd.CellMark);
                        break;
                    case '\u0007': AddControl(DocStoryAtomKind.CellMark, character, cp); break;
                    case '\u001E': AddControl(DocStoryAtomKind.NoBreakHyphen, character, cp); break;
                    case '\u001F': AddControl(DocStoryAtomKind.SoftHyphen, character, cp); break;
                    case '\u0000' when index.CharacterFormatting.Any(x =>
                        x.CpStart <= cp && cp < x.CpEnd && x.Formatting.Special == true):
                        AddControl(DocStoryAtomKind.LegacyPageNumberBlock, character, cp); break;
                    case '\u000F' or '\u0010' or '!' or '%' or '#' or '"'
                        when index.CharacterFormatting.Any(x => x.CpStart <= cp &&
                            cp < x.CpEnd && x.Formatting.Special == true):
                        AddControl(DocStoryAtomKind.LegacyDateBlock, character, cp); break;
                    case '\u0001' when index.CharacterFormatting.Any(x =>
                        x.CpStart <= cp && cp < x.CpEnd && x.Formatting.Special == true &&
                        x.Formatting.PictureDataOffset != null &&
                        x.Formatting.FieldData == true):
                        AddControl(DocStoryAtomKind.FieldData, character, cp); break;
                    case '\u0001' when index.CharacterFormatting.Any(x =>
                        x.CpStart <= cp && cp < x.CpEnd && x.Formatting.Special == true &&
                        x.Formatting.PictureDataOffset != null &&
                        x.Formatting.FieldData != true):
                        AddControl(DocStoryAtomKind.InlinePicture, character, cp); break;
                    case '\u0008': AddControl(DocStoryAtomKind.FloatingShapeAnchor, character, cp); break;
                    case '\u0013' when fieldMarks.Value.TryGetValue(cp, out var field) &&
                        field.Kind == 0x13:
                        AddControl(DocStoryAtomKind.FieldBegin, character, cp); break;
                    case '\u0014' when fieldMarks.Value.TryGetValue(cp, out var field) &&
                        field.Kind == 0x14:
                        AddControl(DocStoryAtomKind.FieldSeparator, character, cp); break;
                    case '\u0015' when fieldMarks.Value.TryGetValue(cp, out var field) &&
                        field.Kind == 0x15:
                        AddControl(DocStoryAtomKind.FieldEnd, character, cp); break;
                    default:
                        if (IsXmlCharacter(character)) AddText(character, cp);
                        else AddControl(DocStoryAtomKind.UnsupportedControl, character, cp);
                        break;
                }
            }
            expected = span.CpEnd;
            // A piece boundary is a potential formatting boundary. Keep it visible
            // to consumers even when adjacent pieces contain identical text.
            if (highSurrogateCp is null) FlushText(expected);
        }
        if (expected != cpEnd)
            throw new InvalidDataException($"The '{partName}' text index does not cover its full CP range.");
        if (highSurrogateCp is uint danglingCp)
            AddControl(DocStoryAtomKind.UnsupportedControl, highSurrogate, danglingCp);
        FlushText(cpEnd);
        if (atoms.Count != 0 || paragraphs.Count == 0)
            paragraphs.Add(new DocStoryParagraph(paragraphStart, cpEnd,
                DocParagraphEnd.None, atoms.ToArray()));
        return new DocStoryText(partName, cpStart, cpEnd, paragraphs);
    }

    private static bool IsXmlCharacter(char value) =>
        value is '\t' or '\n' or '\r' || value >= ' ' &&
        value is not '\uFFFE' and not '\uFFFF';
}
