using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>One-section PlcfSed with an optional Sepx reference.</summary>
internal sealed record DocSectionTable(uint EndCp, int SepxOffset)
{
    public static DocSectionTable Read(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length != 20 || BinaryPrimitives.ReadUInt32LittleEndian(bytes) != 0)
            throw new InvalidDataException("Expected a one-section PlcfSed starting at CP zero.");
        return new DocSectionTable(BinaryPrimitives.ReadUInt32LittleEndian(bytes.Slice(4)),
            BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(10)));
    }

    public byte[] Write()
    {
        if (EndCp == 0 || EndCp >= 0x80000000 || SepxOffset < -1)
            throw new InvalidDataException("The section boundary or Sepx offset is invalid.");
        var bytes = new byte[20];
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4), EndCp);
        BinaryPrimitives.WriteInt32LittleEndian(bytes.AsSpan(10), SepxOffset);
        return bytes;
    }
}
