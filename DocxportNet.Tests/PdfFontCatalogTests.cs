using System.Buffers.Binary;
using System.Text;
using DocxportNet.Visitors.Pdf;

namespace DocxportNet.Tests;

public sealed class PdfFontCatalogTests
{
    [Fact]
    public void KeepsDisplayInLegacyFamilyWhenTypographicStyleContainsIt()
    {
        byte[] font = NameTable((1, "Aptos Display"), (2, "Regular"),
            (16, "Aptos"), (17, "Display"));

        Assert.True(DxpDirectoryPdfFontProvider.TryReadNames(font, out var family, out var style));
        Assert.Equal("Aptos Display", family);
        Assert.Equal("Display", style);
    }

    private static byte[] NameTable(params (ushort Id, string Value)[] names)
    {
        byte[][] strings = names.Select(name => Encoding.BigEndianUnicode.GetBytes(name.Value)).ToArray();
        int tableLength = 6 + names.Length * 12 + strings.Sum(value => value.Length);
        byte[] font = new byte[28 + tableLength];
        BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(0, 4), 0x00010000);
        BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(4, 2), 1);
        Encoding.ASCII.GetBytes("name").CopyTo(font, 12);
        BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(20, 4), 28);
        BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(24, 4), (uint)tableLength);
        BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(30, 2), (ushort)names.Length);
        BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(32, 2), (ushort)(6 + names.Length * 12));
        int stringOffset = 0;
        for (int i = 0; i < names.Length; i++)
        {
            var record = font.AsSpan(34 + i * 12, 12);
            BinaryPrimitives.WriteUInt16BigEndian(record.Slice(0, 2), 3); // Windows Unicode
            BinaryPrimitives.WriteUInt16BigEndian(record.Slice(2, 2), 1);
            BinaryPrimitives.WriteUInt16BigEndian(record.Slice(4, 2), 0x0409);
            BinaryPrimitives.WriteUInt16BigEndian(record.Slice(6, 2), names[i].Id);
            BinaryPrimitives.WriteUInt16BigEndian(record.Slice(8, 2), (ushort)strings[i].Length);
            BinaryPrimitives.WriteUInt16BigEndian(record.Slice(10, 2), (ushort)stringOffset);
            strings[i].CopyTo(font, 34 + names.Length * 12 + stringOffset);
            stringOffset += strings[i].Length;
        }
        return font;
    }
}
