using System.Buffers.Binary;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Doc;

public sealed record DocSectionFormatting(int? Width = null, int? Height = null,
    bool? Landscape = null, int? Left = null, int? Right = null,
    int? Top = null, int? Bottom = null, int? Header = null,
    int? Footer = null, int? Gutter = null, byte? BreakKind = null,
    bool? DifferentFirstPage = null, int? ColumnCount = null, int? ColumnSpace = null,
    IReadOnlyList<(int Width, int Space)>? UnevenColumns = null,
    bool? ColumnSeparator = null, int? PageNumberStart = null,
    byte? VerticalAlignment = null, byte? PageNumberFormat = null,
    DocParagraphBorder? PageBorderTop = null, DocParagraphBorder? PageBorderLeft = null,
    DocParagraphBorder? PageBorderBottom = null, DocParagraphBorder? PageBorderRight = null,
    byte? PageBorderDisplay = null, bool? PageBorderFromPage = null,
    bool? PageBorderBehindText = null, bool? RightToLeft = null,
    bool? GutterOnRight = null,
    ushort? LineNumberCountBy = null, ushort? LineNumberStart = null,
    ushort? LineNumberDistance = null, byte? LineNumberRestart = null,
    ushort? GridLinePitch = null, byte? GridMode = null,
    int? GridCharacterSpace = null)
{
    public bool IsEmpty => this == new DocSectionFormatting();

    // DOC sections need concrete page geometry for independent readers to place
    // their header/footer stories. These values match Word's common Letter setup.
    public DocSectionFormatting WithWriterDefaults() => this with
    {
        Width = Width ?? 12240,
        Height = Height ?? 15840,
        Left = Left ?? 1440,
        Right = Right ?? 1440,
        Top = Top ?? 1440,
        Bottom = Bottom ?? 1440,
        Header = Header ?? 720,
        Footer = Footer ?? 720
    };

    public static DocSectionFormatting FromOpenXml(SectionProperties properties)
    {
        var size = properties.GetFirstChild<PageSize>();
        var margin = properties.GetFirstChild<PageMargin>();
        var columns = properties.GetFirstChild<Columns>();
        var pageBorders = properties.GetFirstChild<PageBorders>();
        var lineNumbers = properties.GetFirstChild<LineNumberType>();
        var grid = properties.GetFirstChild<DocGrid>();
        var breakType = properties.GetFirstChild<SectionType>()?.Val?.Value;
        static int? Number(object? value) => int.TryParse(value?.ToString(), out var n) ? n : null;
        return new DocSectionFormatting(Number(size?.Width), Number(size?.Height),
            size?.Orient == null ? null : size.Orient.Value == PageOrientationValues.Landscape,
            Number(margin?.Left), Number(margin?.Right), Number(margin?.Top),
            Number(margin?.Bottom), Number(margin?.Header), Number(margin?.Footer),
            Number(margin?.Gutter), breakType == null ? null : breakType == SectionMarkValues.Continuous
                ? (byte)0 : breakType == SectionMarkValues.NextColumn ? (byte)1
                : breakType == SectionMarkValues.NextPage ? (byte)2
                : breakType == SectionMarkValues.EvenPage ? (byte)3
                : breakType == SectionMarkValues.OddPage ? (byte)4 : null,
            properties.GetFirstChild<TitlePage>() is { } titlePage
                ? titlePage.Val?.Value ?? true : null,
            Number(columns?.ColumnCount) ?? (columns?.Elements<Column>().Any() == true
                ? columns.Elements<Column>().Count() : null), Number(columns?.Space),
            columns?.EqualWidth?.Value == false
                ? columns.Elements<Column>().Select(column =>
                    (Number(column.Width) ?? 0, Number(column.Space) ?? 0)).ToArray()
                : null,
            columns?.Separator?.Value,
            Number(properties.GetFirstChild<PageNumberType>()?.Start),
            properties.GetFirstChild<VerticalTextAlignmentOnPage>()?.Val?.Value is { } alignment
                ? alignment == VerticalJustificationValues.Center ? (byte)1
                    : alignment == VerticalJustificationValues.Both ? (byte)2
                    : alignment == VerticalJustificationValues.Bottom ? (byte)3 : (byte)0
                : null,
            properties.GetFirstChild<PageNumberType>()?.Format?.Value is { } format
                ? format == NumberFormatValues.Decimal ? (byte)0
                    : format == NumberFormatValues.UpperRoman ? (byte)1
                    : format == NumberFormatValues.LowerRoman ? (byte)2
                    : format == NumberFormatValues.UpperLetter ? (byte)3
                    : format == NumberFormatValues.LowerLetter ? (byte)4 : null
                : null,
            DocParagraphBorder.FromOpenXml(pageBorders?.GetFirstChild<TopBorder>()),
            DocParagraphBorder.FromOpenXml(pageBorders?.GetFirstChild<LeftBorder>()),
            DocParagraphBorder.FromOpenXml(pageBorders?.GetFirstChild<BottomBorder>()),
            DocParagraphBorder.FromOpenXml(pageBorders?.GetFirstChild<RightBorder>()),
            pageBorders?.Display?.Value is { } display
                ? display == PageBorderDisplayValues.FirstPage ? (byte)1
                    : display == PageBorderDisplayValues.NotFirstPage ? (byte)2 : (byte)0
                : null,
            pageBorders?.OffsetFrom?.Value is { } offset
                ? offset == PageBorderOffsetValues.Page : null,
            pageBorders?.ZOrder?.Value is { } order
                ? order == PageBorderZOrderValues.Back : null,
            properties.GetFirstChild<BiDi>() is { } bidi
                ? bidi.Val?.Value ?? true : null,
            properties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.GutterOnRight>() is { } rtlGutter
                ? rtlGutter.Val?.Value ?? true : null,
            lineNumbers == null ? null : checked((ushort)(lineNumbers.CountBy?.Value ?? 1)),
            lineNumbers?.Start?.Value is short lineStart && lineStart >= 0
                ? checked((ushort)lineStart) : null,
            Number(lineNumbers?.Distance) is int lineDistance && lineDistance >= 0
                ? checked((ushort)lineDistance) : null,
            lineNumbers?.Restart?.Value is { } restart
                ? restart == LineNumberRestartValues.NewSection ? (byte)1
                    : restart == LineNumberRestartValues.Continuous ? (byte)2
                    : (byte)0 : null,
            grid != null && (grid.Type?.Value == DocGridValues.Lines ||
                grid.Type?.Value == DocGridValues.LinesAndChars ||
                grid.Type?.Value == DocGridValues.SnapToChars) &&
                Number(grid.LinePitch) is int pitch ? checked((ushort)pitch) : null,
            grid?.Type?.Value == DocGridValues.LinesAndChars ? (byte)1
                : grid?.Type?.Value == DocGridValues.Lines ? (byte)2
                : grid?.Type?.Value == DocGridValues.SnapToChars ? (byte)3 : null,
            grid?.Type?.Value == DocGridValues.LinesAndChars ||
                grid?.Type?.Value == DocGridValues.SnapToChars
                ? Number(grid.CharacterSpace) : null);
    }

    public void ApplyTo(SectionProperties properties)
    {
        if (BreakKind is byte kind)
            properties.AppendChild(new SectionType { Val = kind switch
            {
                0 => SectionMarkValues.Continuous,
                1 => SectionMarkValues.NextColumn,
                3 => SectionMarkValues.EvenPage,
                4 => SectionMarkValues.OddPage,
                _ => SectionMarkValues.NextPage
            } });
        if (Width != null || Height != null || Landscape != null)
        {
            var size = new PageSize();
            if (Width is int width) size.Width = (uint)width;
            if (Height is int height) size.Height = (uint)height;
            if (Landscape is bool landscape)
                size.Orient = landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait;
            properties.AppendChild(size);
        }
        if (Left != null || Right != null || Top != null || Bottom != null ||
            Header != null || Footer != null || Gutter != null)
        {
            var margin = new PageMargin();
            if (Left is int left) margin.Left = (uint)left;
            if (Right is int right) margin.Right = (uint)right;
            if (Top is int top) margin.Top = top;
            if (Bottom is int bottom) margin.Bottom = bottom;
            if (Header is int header) margin.Header = (uint)header;
            if (Footer is int footer) margin.Footer = (uint)footer;
            if (Gutter is int gutter) margin.Gutter = (uint)gutter;
            properties.AppendChild(margin);
        }
        if (PageBorderTop != null || PageBorderLeft != null ||
            PageBorderBottom != null || PageBorderRight != null ||
            PageBorderDisplay != null || PageBorderFromPage != null ||
            PageBorderBehindText != null)
        {
            var borders = new PageBorders();
            if (PageBorderDisplay is byte display)
                borders.Display = display switch
                {
                    1 => PageBorderDisplayValues.FirstPage,
                    2 => PageBorderDisplayValues.NotFirstPage,
                    _ => PageBorderDisplayValues.AllPages
                };
            if (PageBorderFromPage is bool fromPage)
                borders.OffsetFrom = fromPage ? PageBorderOffsetValues.Page : PageBorderOffsetValues.Text;
            if (PageBorderBehindText is bool behind)
                borders.ZOrder = behind ? PageBorderZOrderValues.Back : PageBorderZOrderValues.Front;
            if (PageBorderTop is { } top)
            { var edge = new TopBorder(); top.ApplyTo(edge); borders.AppendChild(edge); }
            if (PageBorderLeft is { } left)
            { var edge = new LeftBorder(); left.ApplyTo(edge); borders.AppendChild(edge); }
            if (PageBorderBottom is { } bottom)
            { var edge = new BottomBorder(); bottom.ApplyTo(edge); borders.AppendChild(edge); }
            if (PageBorderRight is { } right)
            { var edge = new RightBorder(); right.ApplyTo(edge); borders.AppendChild(edge); }
            properties.AppendChild(borders);
        }
        if (PageNumberStart != null || PageNumberFormat != null)
        {
            var numbering = new PageNumberType();
            if (PageNumberStart is int pageStart) numbering.Start = pageStart;
            if (PageNumberFormat is byte pageFormat)
                numbering.Format = pageFormat switch
                {
                    1 => NumberFormatValues.UpperRoman,
                    2 => NumberFormatValues.LowerRoman,
                    3 => NumberFormatValues.UpperLetter,
                    4 => NumberFormatValues.LowerLetter,
                    _ => NumberFormatValues.Decimal
                };
            properties.AppendChild(numbering);
        }
        if (ColumnCount != null || ColumnSpace != null || UnevenColumns != null ||
            ColumnSeparator != null)
        {
            var columns = new Columns();
            if (ColumnCount is int count) columns.ColumnCount = (short)count;
            if (ColumnSpace is int space) columns.Space = space.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (ColumnSeparator is bool separator) columns.Separator = separator;
            if (UnevenColumns is { Count: > 0 } uneven)
            {
                columns.EqualWidth = false;
                foreach (var (width, gap) in uneven)
                    columns.AppendChild(new Column
                    {
                        Width = width.ToString(System.Globalization.CultureInfo.InvariantCulture),
                        Space = gap.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    });
            }
            properties.AppendChild(columns);
        }
        if (VerticalAlignment is byte verticalAlignment)
            properties.AppendChild(new VerticalTextAlignmentOnPage
            {
                Val = verticalAlignment switch
                {
                    1 => VerticalJustificationValues.Center,
                    2 => VerticalJustificationValues.Both,
                    3 => VerticalJustificationValues.Bottom,
                    _ => VerticalJustificationValues.Top
                }
            });
        if (LineNumberCountBy is > 0 and <= 100)
        {
            var lineNumbers = new LineNumberType
            {
                CountBy = checked((short)LineNumberCountBy.Value),
                Restart = LineNumberRestart switch
                {
                    1 => LineNumberRestartValues.NewSection,
                    2 => LineNumberRestartValues.Continuous,
                    _ => LineNumberRestartValues.NewPage
                }
            };
            if (LineNumberStart is ushort start)
                lineNumbers.Start = checked((short)start);
            if (LineNumberDistance is ushort distance)
                lineNumbers.Distance = distance.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            properties.AddChild(lineNumbers, true);
        }
        if (GridLinePitch is ushort pitch)
            properties.AddChild(new DocGrid
            {
                Type = GridMode switch
                {
                    1 => DocGridValues.LinesAndChars,
                    3 => DocGridValues.SnapToChars,
                    _ => DocGridValues.Lines
                },
                LinePitch = pitch,
                CharacterSpace = GridMode is 1 or 3 ? GridCharacterSpace : null
            }, true);
        if (DifferentFirstPage == true) properties.AppendChild(new TitlePage());
        if (RightToLeft is bool rightToLeft)
            properties.AppendChild(new BiDi { Val = rightToLeft });
        if (GutterOnRight is bool gutterOnRight)
            properties.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.GutterOnRight
                { Val = gutterOnRight });
    }

    public byte[] Encode()
    {
        var bytes = new List<byte>();
        void Word(ushort sprm, int? value)
        {
            if (value is not int n) return;
            if (n < short.MinValue || n > ushort.MaxValue)
                throw new InvalidDataException("A section measurement exceeds the DOC twip range.");
            bytes.Add((byte)sprm); bytes.Add((byte)(sprm >> 8));
            bytes.Add((byte)n); bytes.Add((byte)(n >> 8));
        }
        if (Landscape is bool landscape)
            bytes.AddRange([0x1D, 0x30, (byte)(landscape ? 2 : 1)]);
        if (ColumnSeparator is bool separator)
            bytes.AddRange([0x19, 0x30, (byte)(separator ? 1 : 0)]);
        if (VerticalAlignment is byte verticalAlignment)
            bytes.AddRange([0x1A, 0x30, verticalAlignment]);
        if (LineNumberCountBy is ushort countBy)
        {
            if (countBy > 100)
                throw new InvalidDataException("A DOC line-number interval is invalid.");
            bytes.AddRange([0x13, 0x30, LineNumberRestart ?? 0]);
            Word(0x5015, countBy);
            Word(0x9016, LineNumberDistance ?? 0);
            Word(0x501B, LineNumberStart ?? 0);
        }
        if (GridLinePitch is ushort gridPitch)
        {
            if (gridPitch == 0 || gridPitch > 31680)
                throw new InvalidDataException("A DOC grid line pitch is invalid.");
            if (GridMode is not (null or 1 or 2 or 3))
                throw new InvalidDataException("A DOC grid mode is invalid.");
            if (GridCharacterSpace is int charSpace && GridMode is 1 or 3)
            {
                if (charSpace is < -670925 or > 6488064)
                    throw new InvalidDataException("A DOC grid character pitch is invalid.");
                bytes.AddRange([0x30, 0x70, (byte)charSpace, (byte)(charSpace >> 8),
                    (byte)(charSpace >> 16), (byte)(charSpace >> 24)]);
            }
            Word(0x9031, gridPitch);
            Word(0x5032, GridMode ?? 2);
        }
        if (RightToLeft is bool rightToLeft)
            bytes.AddRange([0x28, 0x32, (byte)(rightToLeft ? 1 : 0)]);
        if (GutterOnRight is bool gutterOnRight)
            bytes.AddRange([0x2A, 0x32, (byte)(gutterOnRight ? 1 : 0)]);
        if (BreakKind is byte kind) bytes.AddRange([0x09, 0x30, kind]);
        if (DifferentFirstPage is bool first)
            bytes.AddRange([0x0A, 0x30, (byte)(first ? 1 : 0)]);
        if (PageNumberStart is int pageStart)
        {
            if (pageStart < 0 || pageStart > 32766)
                throw new InvalidDataException("A section page number exceeds the DOC range.");
            bytes.AddRange([0x11, 0x30, 1]);
            Word(0x501C, pageStart);
        }
        if (PageNumberFormat is byte pageFormat)
            bytes.AddRange([0x0E, 0x30, pageFormat]);
        Word(0x500B, ColumnCount - 1);
        Word(0x900C, ColumnSpace);
        if (UnevenColumns is { Count: > 0 } uneven)
        {
            bytes.AddRange([0x05, 0x30, 0]);
            for (var column = 0; column < uneven.Count; column++)
            {
                void ColumnOperand(ushort sprm, int value)
                {
                    if (value < 0 || value > ushort.MaxValue)
                        throw new InvalidDataException("A section column measurement exceeds the DOC twip range.");
                    bytes.Add((byte)sprm); bytes.Add((byte)(sprm >> 8));
                    bytes.Add((byte)column); bytes.Add((byte)value); bytes.Add((byte)(value >> 8));
                }
                ColumnOperand(0xF203, uneven[column].Width);
                if (column + 1 < uneven.Count)
                    ColumnOperand(0xF204, uneven[column].Space);
            }
        }
        Word(0xB01F, Width); Word(0xB020, Height);
        Word(0xB021, Left); Word(0xB022, Right);
        Word(0x9023, Top); Word(0x9024, Bottom);
        Word(0xB017, Header); Word(0xB018, Footer); Word(0xB025, Gutter);
        void PageBorder(ushort legacySprm, ushort modernSprm, DocParagraphBorder? border)
        {
            if (border == null) return;
            if (border.Encode80() is { } legacy)
            {
                bytes.Add((byte)legacySprm); bytes.Add((byte)(legacySprm >> 8));
                bytes.AddRange(legacy);
            }
            bytes.Add((byte)modernSprm); bytes.Add((byte)(modernSprm >> 8));
            bytes.AddRange(border.Encode());
        }
        PageBorder(0x702B, 0xD234, PageBorderTop);
        PageBorder(0x702C, 0xD235, PageBorderLeft);
        PageBorder(0x702D, 0xD236, PageBorderBottom);
        PageBorder(0x702E, 0xD237, PageBorderRight);
        if (PageBorderDisplay != null || PageBorderFromPage != null ||
            PageBorderBehindText != null)
        {
            var display = PageBorderDisplay ?? 0;
            if (display > 2) throw new InvalidDataException("Invalid DOC page-border display setting.");
            bytes.AddRange([0x2F, 0x52,
                (byte)(display | (PageBorderBehindText == true ? 0x08 : 0) |
                    (PageBorderFromPage == true ? 0x20 : 0)), 0]);
        }
        return bytes.ToArray();
    }

    public static DocSectionFormatting Read(DocTextIndex index, DocStructureNode section)
    {
        var sepx = section.Children.FirstOrDefault(x => x.Kind == "Sepx");
        if (sepx?.Offset == null || sepx.Length == null) return new DocSectionFormatting();
        var bytes = index.Structure.ReadRange("WordDocument", sepx.Offset.Value + 2,
            checked((int)sepx.Length.Value - 2));
        int? width = null, height = null, left = null, right = null, top = null,
            bottom = null, header = null, footer = null, gutter = null;
        bool? landscape = null, firstPage = null, pageNumberRestart = null;
        byte? breakKind = null, verticalAlignment = null, pageNumberFormat = null;
        ushort? lineNumberCountBy = null, lineNumberStart = null,
            lineNumberDistance = null;
        byte? lineNumberRestart = null;
        ushort? gridPitch = null;
        byte? gridMode = null;
        int? gridCharacterSpace = null;
        int? columnCount = null, columnSpace = null;
        bool? evenlySpaced = null, columnSeparator = null;
        int? pageNumberStart97 = null, pageNumberStartModern = null;
        var widths = new Dictionary<int, int>();
        var spaces = new Dictionary<int, int>();
        DocParagraphBorder? pageBorderTop = null, pageBorderLeft = null,
            pageBorderBottom = null, pageBorderRight = null;
        byte? pageBorderDisplay = null;
        bool? pageBorderFromPage = null, pageBorderBehindText = null;
        bool? rightToLeft = null, gutterOnRight = null;
        for (var i = 0; i + 2 <= bytes.Length;)
        {
            var sprm = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(i)); i += 2;
            var spra = sprm >> 13;
            var length = spra switch { 0 or 1 => 1, 2 or 4 or 5 => 2,
                3 => 4, 7 => 3, 6 => i < bytes.Length ? bytes[i] + 1 : 0, _ => 0 };
            if (length == 0 || i + length > bytes.Length) break;
            var unsigned = length >= 2 ? BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(i)) : 0;
            var signed = length >= 2 ? (short)unsigned : (short)0;
            switch (sprm)
            {
                case 0x301D: landscape = bytes[i] == 2; break;
                case 0x3009: breakKind = bytes[i]; break;
                case 0x300A: firstPage = bytes[i] != 0; break;
                case 0x3011: pageNumberRestart = bytes[i] != 0; break;
                case 0x300E: pageNumberFormat = bytes[i] <= 4 ? bytes[i] : null; break;
                case 0x501C: pageNumberStart97 = unsigned; break;
                case 0x7044: pageNumberStartModern = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i))); break;
                case 0x500B: columnCount = unsigned + 1; break;
                case 0x900C: columnSpace = unsigned; break;
                case 0x3005: evenlySpaced = bytes[i] != 0; break;
                case 0x3019: columnSeparator = bytes[i] != 0; break;
                case 0x301A: verticalAlignment = bytes[i]; break;
                case 0x3013: lineNumberRestart = bytes[i] <= 2 ? bytes[i] : null; break;
                case 0x5015: lineNumberCountBy = checked((ushort)unsigned); break;
                case 0x9016: lineNumberDistance = checked((ushort)unsigned); break;
                case 0x501B: lineNumberStart = checked((ushort)unsigned); break;
                case 0x9031: gridPitch = checked((ushort)unsigned); break;
                case 0x5032: gridMode = unsigned is >= 1 and <= 3
                    ? checked((byte)unsigned) : null; break;
                case 0x7030: gridCharacterSpace =
                    BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(i)); break;
                case 0x3228: rightToLeft = bytes[i] != 0; break;
                case 0x322A: gutterOnRight = bytes[i] != 0; break;
                case 0xF203: widths[bytes[i]] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(i + 1)); break;
                case 0xF204: spaces[bytes[i]] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(i + 1)); break;
                case 0xB01F: width = unsigned; break;
                case 0xB020: height = unsigned; break;
                case 0xB021: left = unsigned; break;
                case 0xB022: right = unsigned; break;
                case 0x9023: top = signed; break;
                case 0x9024: bottom = signed; break;
                case 0xB017: header = unsigned; break;
                case 0xB018: footer = unsigned; break;
                case 0xB025: gutter = unsigned; break;
                case 0x702B: pageBorderTop = DocParagraphBorder.Parse80(bytes.AsSpan(i, length)); break;
                case 0x702C: pageBorderLeft = DocParagraphBorder.Parse80(bytes.AsSpan(i, length)); break;
                case 0x702D: pageBorderBottom = DocParagraphBorder.Parse80(bytes.AsSpan(i, length)); break;
                case 0x702E: pageBorderRight = DocParagraphBorder.Parse80(bytes.AsSpan(i, length)); break;
                case 0xD234: pageBorderTop = DocParagraphBorder.Parse(bytes.AsSpan(i, length)); break;
                case 0xD235: pageBorderLeft = DocParagraphBorder.Parse(bytes.AsSpan(i, length)); break;
                case 0xD236: pageBorderBottom = DocParagraphBorder.Parse(bytes.AsSpan(i, length)); break;
                case 0xD237: pageBorderRight = DocParagraphBorder.Parse(bytes.AsSpan(i, length)); break;
                case 0x522F:
                    pageBorderDisplay = (byte)(bytes[i] & 7);
                    pageBorderBehindText = (bytes[i] & 8) != 0;
                    pageBorderFromPage = (bytes[i] & 0x20) != 0;
                    break;
            }
            i += length;
        }
        IReadOnlyList<(int Width, int Space)>? unevenColumns = null;
        if (evenlySpaced == false && columnCount is > 0 &&
            Enumerable.Range(0, columnCount.Value).All(widths.ContainsKey))
            unevenColumns = Enumerable.Range(0, columnCount.Value)
                .Select(column => (widths[column], spaces.TryGetValue(column, out var space) ? space : 0))
                .ToArray();
        return new DocSectionFormatting(width, height, landscape, left, right,
            top, bottom, header, footer, gutter, breakKind, firstPage,
            columnCount, columnSpace, unevenColumns, columnSeparator,
            pageNumberRestart == true ? pageNumberStartModern ?? pageNumberStart97 ?? 0 : null,
            verticalAlignment, pageNumberFormat,
            pageBorderTop, pageBorderLeft, pageBorderBottom, pageBorderRight,
            pageBorderDisplay, pageBorderFromPage, pageBorderBehindText,
            rightToLeft, gutterOnRight, lineNumberCountBy,
            lineNumberStart, lineNumberDistance, lineNumberRestart,
            gridMode != null ? gridPitch : null, gridMode,
            gridMode is 1 or 3 ? gridCharacterSpace : null);
    }
}
