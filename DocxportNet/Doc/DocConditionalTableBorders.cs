using System.Buffers.Binary;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Doc;

public sealed record DocConditionalTableBorders(
    DocParagraphBorder? Top = null, DocParagraphBorder? Bottom = null,
    DocParagraphBorder? Left = null, DocParagraphBorder? Right = null,
    DocParagraphBorder? InsideHorizontal = null,
    DocParagraphBorder? InsideVertical = null,
    DocParagraphBorder? TopLeftToBottomRight = null,
    DocParagraphBorder? TopRightToBottomLeft = null)
{
    internal static DocConditionalTableBorders? FromOpenXml(TableCellBorders? source)
    {
        if (source == null) return null;
        var result = new DocConditionalTableBorders(
            DocParagraphBorder.FromOpenXml(source.TopBorder),
            DocParagraphBorder.FromOpenXml(source.BottomBorder),
            DocParagraphBorder.FromOpenXml((BorderType?)source.StartBorder ?? source.LeftBorder),
            DocParagraphBorder.FromOpenXml((BorderType?)source.EndBorder ?? source.RightBorder),
            DocParagraphBorder.FromOpenXml(source.InsideHorizontalBorder),
            DocParagraphBorder.FromOpenXml(source.InsideVerticalBorder),
            DocParagraphBorder.FromOpenXml(source.TopLeftToBottomRightCellBorder),
            DocParagraphBorder.FromOpenXml(source.TopRightToBottomLeftCellBorder));
        return result.IsEmpty ? null : result;
    }

    internal bool IsEmpty => Top == null && Bottom == null && Left == null &&
        Right == null && InsideHorizontal == null && InsideVertical == null &&
        TopLeftToBottomRight == null && TopRightToBottomLeft == null;

    internal byte[] Encode()
    {
        using var stream = new MemoryStream();
        void Write(ushort code, DocParagraphBorder? border)
        {
            if (border == null) return;
            Span<byte> opcode = stackalloc byte[2];
            BinaryPrimitives.WriteUInt16LittleEndian(opcode, code);
            stream.Write(opcode);
            stream.Write(border.Encode());
        }
        Write(0xD47F, Top);
        Write(0xD680, Bottom);
        Write(0xD681, Left);
        Write(0xD682, Right);
        Write(0xD683, InsideHorizontal);
        Write(0xD684, InsideVertical);
        Write(0xD685, TopLeftToBottomRight);
        Write(0xD686, TopRightToBottomLeft);
        return stream.ToArray();
    }

    internal static DocConditionalTableBorders? Parse(ReadOnlySpan<byte> sprms)
    {
        var result = new DocConditionalTableBorders();
        for (var offset = 0; offset < sprms.Length;)
        {
            if (sprms.Length - offset < 2)
                throw new InvalidDataException("A conditional table border rule is truncated.");
            var code = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
            var operandOffset = offset + 2;
            var count = (code >> 13) switch
            {
                0 or 1 => 1,
                2 or 4 or 5 => 2,
                3 => 4,
                6 when operandOffset < sprms.Length => 1 + sprms[operandOffset],
                7 => 3,
                _ => throw new InvalidDataException("A conditional table border rule has an invalid operand.")
            };
            if (sprms.Length - operandOffset < count)
                throw new InvalidDataException("A conditional table border rule has an invalid length.");
            if (code is 0xD47F or 0xD680 or 0xD681 or 0xD682 or 0xD683 or
                0xD684 or 0xD685 or 0xD686)
            {
                var border = DocParagraphBorder.Parse(sprms.Slice(operandOffset,
                    count));
                result = code switch
                {
                    0xD47F => result with { Top = border },
                    0xD680 => result with { Bottom = border },
                    0xD681 => result with { Left = border },
                    0xD682 => result with { Right = border },
                    0xD683 => result with { InsideHorizontal = border },
                    0xD684 => result with { InsideVertical = border },
                    0xD685 => result with { TopLeftToBottomRight = border },
                    0xD686 => result with { TopRightToBottomLeft = border },
                    _ => result
                };
            }
            offset = operandOffset + count;
        }
        return result.IsEmpty ? null : result;
    }

    internal TableCellBorders ToOpenXml()
    {
        var result = new TableCellBorders();
        if (Top is { } top) { var edge = new TopBorder(); top.ApplyTo(edge); result.AppendChild(edge); }
        if (Left is { } left) { var edge = new LeftBorder(); left.ApplyTo(edge); result.AppendChild(edge); }
        if (Bottom is { } bottom) { var edge = new BottomBorder(); bottom.ApplyTo(edge); result.AppendChild(edge); }
        if (Right is { } right) { var edge = new RightBorder(); right.ApplyTo(edge); result.AppendChild(edge); }
        if (InsideHorizontal is { } insideH) { var edge = new InsideHorizontalBorder(); insideH.ApplyTo(edge); result.AppendChild(edge); }
        if (InsideVertical is { } insideV) { var edge = new InsideVerticalBorder(); insideV.ApplyTo(edge); result.AppendChild(edge); }
        if (TopLeftToBottomRight is { } down) { var edge = new TopLeftToBottomRightCellBorder(); down.ApplyTo(edge); result.AppendChild(edge); }
        if (TopRightToBottomLeft is { } up) { var edge = new TopRightToBottomLeftCellBorder(); up.ApplyTo(edge); result.AppendChild(edge); }
        return result;
    }
}
