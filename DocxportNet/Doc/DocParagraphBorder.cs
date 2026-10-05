using System.Buffers.Binary;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxportNet.Doc;

/// <summary>Selected BRC fields for a paragraph edge.</summary>
public sealed record DocParagraphBorder(byte Type, byte WidthEighthPoints,
    byte SpacePoints, uint? ColorRgb, bool Shadow = false, bool Frame = false)
{
    internal static DocParagraphBorder? Parse(ReadOnlySpan<byte> operand)
    {
        if (operand.Length != 9 || operand[0] != 8) return null;
        var type = operand[6];
        if (ToOpenXmlType(type) == null) return null;
        var color = BinaryPrimitives.ReadUInt32LittleEndian(operand.Slice(1, 4));
        var flags = BinaryPrimitives.ReadUInt16LittleEndian(operand.Slice(7, 2));
        return new DocParagraphBorder(type, operand[5], (byte)(flags & 31),
            (color >> 24) == 0 ? color & 0xFFFFFF : null,
            (flags & 32) != 0, (flags & 64) != 0);
    }

    internal static DocParagraphBorder? ParseRaw(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length != 8) return null;
        Span<byte> operand = stackalloc byte[9];
        operand[0] = 8;
        bytes.CopyTo(operand.Slice(1));
        return Parse(operand);
    }

    internal static DocParagraphBorder? Parse80(ReadOnlySpan<byte> operand)
    {
        if (operand.Length != 4 || ToOpenXmlType(operand[1]) == null ||
            operand[2] > 16) return null;
        uint? color = operand[2] switch
        {
            0 => null,
            1 => 0x000000u,
            2 => 0xFF0000u,
            3 => 0xFFFF00u,
            4 => 0x00FF00u,
            5 => 0xFF00FFu,
            6 => 0x0000FFu,
            7 => 0x00FFFFu,
            8 => 0xFFFFFFu,
            9 => 0x800000u,
            10 => 0x808000u,
            11 => 0x008000u,
            12 or 13 => 0x800080u,
            14 => 0x008080u,
            15 => 0x808080u,
            16 => 0xC0C0C0u,
            _ => null
        };
        var flags = operand[3];
        return new DocParagraphBorder(operand[1], operand[0],
            (byte)(flags & 31), color, (flags & 32) != 0, (flags & 64) != 0);
    }

    internal byte[] EncodeRaw() => Encode().AsSpan(1).ToArray();

    internal byte[]? Encode80()
    {
        if (Type is 26 or 27) return null;
        if (SpacePoints > 31)
            throw new InvalidDataException("DOC border spacing exceeds 31 points.");
        byte color = ColorRgb switch
        {
            null => 0,
            0x000000u => 1, 0xFF0000u => 2,
            0xFFFF00u => 3, 0x00FF00u => 4,
            0xFF00FFu => 5, 0x0000FFu => 6,
            0x00FFFFu => 7, 0xFFFFFFu => 8,
            0x800000u => 9, 0x808000u => 10,
            0x008000u => 11, 0x800080u => 12,
            0x008080u => 14, 0x808080u => 15,
            0xC0C0C0u => 16, _ => 0
        };
        return new[] { WidthEighthPoints, Type, color,
            (byte)(SpacePoints | (Shadow ? 32 : 0) | (Frame ? 64 : 0)) };
    }

    internal byte[] Encode()
    {
        if (SpacePoints > 31) throw new InvalidDataException("DOC border spacing exceeds 31 points.");
        var bytes = new byte[9];
        bytes[0] = 8;
        var color = ColorRgb ?? 0xFF000000u;
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(1), color);
        bytes[5] = WidthEighthPoints;
        bytes[6] = Type;
        var flags = (ushort)(SpacePoints | (Shadow ? 32 : 0) | (Frame ? 64 : 0));
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(7), flags);
        return bytes;
    }

    internal static DocParagraphBorder? FromOpenXml(BorderType? border)
    {
        if (border?.Val?.Value is not { } value) return null;
        var type = FromOpenXmlType(value);
        if (type == null) return null;
        var width = byte.TryParse(border.Size?.Value.ToString(), out var size) ? size : (byte)4;
        var space = byte.TryParse(border.Space?.Value.ToString(), out var parsedSpace)
            ? parsedSpace : (byte)0;
        uint? color = null;
        var raw = border.Color?.Value;
        if (raw?.Length == 6 && uint.TryParse(raw,
            System.Globalization.NumberStyles.HexNumber,
            System.Globalization.CultureInfo.InvariantCulture, out var rgb))
            color = ((rgb & 0xFF) << 16) | (rgb & 0x00FF00) | (rgb >> 16);
        return new DocParagraphBorder(type.Value, width, space, color,
            border.Shadow?.Value ?? false, border.Frame?.Value ?? false);
    }

    internal void ApplyTo(BorderType border)
    {
        border.Val = ToOpenXmlType(Type);
        border.Size = WidthEighthPoints;
        border.Space = SpacePoints;
        if (ColorRgb is uint color)
        {
            var rgb = ((color & 0xFF) << 16) | (color & 0x00FF00) | (color >> 16);
            border.Color = rgb.ToString("X6", System.Globalization.CultureInfo.InvariantCulture);
        }
        else border.Color = "auto";
        if (Shadow) border.Shadow = true;
        if (Frame) border.Frame = true;
    }

    private static byte? FromOpenXmlType(BorderValues value) =>
        value == BorderValues.None || value == BorderValues.Nil ? (byte)0 :
        value == BorderValues.Single ? (byte)1 :
        value == BorderValues.Double ? (byte)3 :
        value == BorderValues.Dotted ? (byte)6 :
        value == BorderValues.Dashed ? (byte)7 :
        value == BorderValues.DotDash ? (byte)8 :
        value == BorderValues.DotDotDash ? (byte)9 :
        value == BorderValues.Triple ? (byte)10 :
        value == BorderValues.ThinThickSmallGap ? (byte)11 :
        value == BorderValues.ThickThinSmallGap ? (byte)12 :
        value == BorderValues.ThinThickThinSmallGap ? (byte)13 :
        value == BorderValues.ThinThickMediumGap ? (byte)14 :
        value == BorderValues.ThickThinMediumGap ? (byte)15 :
        value == BorderValues.ThinThickThinMediumGap ? (byte)16 :
        value == BorderValues.ThinThickLargeGap ? (byte)17 :
        value == BorderValues.ThickThinLargeGap ? (byte)18 :
        value == BorderValues.ThinThickThinLargeGap ? (byte)19 :
        value == BorderValues.Wave ? (byte)20 :
        value == BorderValues.DoubleWave ? (byte)21 :
        value == BorderValues.DashSmallGap ? (byte)22 :
        value == BorderValues.DashDotStroked ? (byte)23 :
        value == BorderValues.ThreeDEmboss ? (byte)24 :
        value == BorderValues.ThreeDEngrave ? (byte)25 :
        value == BorderValues.Outset ? (byte)26 :
        value == BorderValues.Inset ? (byte)27 : null;

    private static BorderValues? ToOpenXmlType(byte value) => value switch
    {
        0 => BorderValues.None, 1 or 5 => BorderValues.Single,
        3 => BorderValues.Double, 6 => BorderValues.Dotted,
        7 => BorderValues.Dashed, 8 => BorderValues.DotDash,
        9 => BorderValues.DotDotDash, 10 => BorderValues.Triple,
        11 => BorderValues.ThinThickSmallGap,
        12 => BorderValues.ThickThinSmallGap,
        13 => BorderValues.ThinThickThinSmallGap,
        14 => BorderValues.ThinThickMediumGap,
        15 => BorderValues.ThickThinMediumGap,
        16 => BorderValues.ThinThickThinMediumGap,
        17 => BorderValues.ThinThickLargeGap,
        18 => BorderValues.ThickThinLargeGap,
        19 => BorderValues.ThinThickThinLargeGap,
        20 => BorderValues.Wave, 21 => BorderValues.DoubleWave,
        22 => BorderValues.DashSmallGap, 23 => BorderValues.DashDotStroked,
        24 => BorderValues.ThreeDEmboss, 25 => BorderValues.ThreeDEngrave,
        26 => BorderValues.Outset, 27 => BorderValues.Inset, _ => null
    };
}
