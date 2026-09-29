using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>Writes the Word 2002 FIB directory and its table-block locations.</summary>
internal sealed class DocFibWriter
{
    private const int PairOffset = 154;
    private readonly byte[] word;

    public DocFibWriter(byte[] word, int textStart, int textEnd, int characterCount)
    {
        this.word = word;
        U16(0, 0xA5EC); // wIdent
        U16(2, 0x0101); // nFib: Word 2002
        U16(4, 0x204D); // unused, conventional value
        U16(6, 0x0409); // English language ID
        U16(10, 0x12F0); // Unicode, 1Table, complex piece table, quick-save count
        U16(12, 0x00BF); // nFibBack
        U32(24, textStart);
        U32(28, textEnd);
        U16(32, 14); // csw
        U16(62, 22); // cslw
        U32(66, word.Length); // cbMac
        U32(76, characterCount); // ccpText
        U16(152, 136); // cbRgFcLcb (Word 2002)
        U16(PairOffset + 136 * 8, 0); // cswNew
    }

    public void AddTableBlock(MemoryStream table, int pairIndex, byte[] bytes)
    {
        var offset = checked((int)table.Position);
        table.Write(bytes, 0, bytes.Length);
        U32(PairOffset + pairIndex * 8, offset);
        U32(PairOffset + pairIndex * 8 + 4, bytes.Length);
    }

    private void U16(int offset, int value) =>
        BinaryPrimitives.WriteUInt16LittleEndian(word.AsSpan(offset), checked((ushort)value));
    private void U32(int offset, int value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(word.AsSpan(offset), checked((uint)value));
}
