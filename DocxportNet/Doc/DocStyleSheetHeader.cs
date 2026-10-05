using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>The fixed STSHI fields needed to locate and write fixed-index styles.</summary>
internal sealed record DocStyleSheetHeader(ushort StyleCount, ushort BaseSize,
    ushort Flags, ushort MaxBuiltInStyle, ushort FixedStyleCount,
    short AsciiFontIndex = 0, short EastAsiaFontIndex = 0,
    short HighAnsiFontIndex = 0, short ComplexScriptFontIndex = 0)
{
    public static DocStyleSheetHeader Read(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 20) throw new InvalidDataException("The stylesheet header is truncated.");
        var length = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        if (length < 18)
            throw new InvalidDataException("The stylesheet header length is invalid.");
        return new DocStyleSheetHeader(
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(2)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(4)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(10)),
            BinaryPrimitives.ReadInt16LittleEndian(bytes.Slice(14)),
            BinaryPrimitives.ReadInt16LittleEndian(bytes.Slice(16)),
            BinaryPrimitives.ReadInt16LittleEndian(bytes.Slice(18)),
            length >= 20 && bytes.Length >= 22
                ? BinaryPrimitives.ReadInt16LittleEndian(bytes.Slice(20)) : (short)0);
    }

    public byte[] Write()
    {
        if (StyleCount < 15 || StyleCount >= 0x0FFE || FixedStyleCount != 15)
            throw new InvalidDataException("The stylesheet requires fifteen fixed-index style slots.");
        // StshiLsd is present even with no latent styles: cbLSD=4 and an
        // empty mpstiilsd when stiMaxWhenSaved is zero.
        var bytes = new byte[24];
        Put(bytes, 0, 22);
        Put(bytes, 2, StyleCount);
        Put(bytes, 4, BaseSize);
        Put(bytes, 6, Flags);
        Put(bytes, 8, MaxBuiltInStyle);
        Put(bytes, 10, FixedStyleCount);
        BinaryPrimitives.WriteInt16LittleEndian(bytes.AsSpan(14), AsciiFontIndex);
        BinaryPrimitives.WriteInt16LittleEndian(bytes.AsSpan(16), EastAsiaFontIndex);
        BinaryPrimitives.WriteInt16LittleEndian(bytes.AsSpan(18), HighAnsiFontIndex);
        BinaryPrimitives.WriteInt16LittleEndian(bytes.AsSpan(20), ComplexScriptFontIndex);
        Put(bytes, 22, 4);
        return bytes;
    }

    private static void Put(byte[] bytes, int offset, ushort value) =>
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset), value);
}
