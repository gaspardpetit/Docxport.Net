using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>The fixed STSHI fields needed to locate and write fixed-index styles.</summary>
internal sealed record DocStyleSheetHeader(ushort StyleCount, ushort BaseSize,
    ushort Flags, ushort MaxBuiltInStyle, ushort FixedStyleCount)
{
    public static DocStyleSheetHeader Read(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 20) throw new InvalidDataException("The stylesheet header is truncated.");
        var length = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        if (length < 18 || bytes.Length < length + 2)
            throw new InvalidDataException("The stylesheet header length is invalid.");
        return new DocStyleSheetHeader(
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(2)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(4)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8)),
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(10)));
    }

    public byte[] Write()
    {
        if (StyleCount < 15 || StyleCount >= 0x0FFE || FixedStyleCount != 15)
            throw new InvalidDataException("The stylesheet requires fifteen fixed-index style slots.");
        var bytes = new byte[20];
        Put(bytes, 0, 18);
        Put(bytes, 2, StyleCount);
        Put(bytes, 4, BaseSize);
        Put(bytes, 6, Flags);
        Put(bytes, 8, MaxBuiltInStyle);
        Put(bytes, 10, FixedStyleCount);
        return bytes;
    }

    private static void Put(byte[] bytes, int offset, ushort value) =>
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset), value);
}
