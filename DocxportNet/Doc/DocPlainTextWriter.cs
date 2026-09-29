using System.Buffers.Binary;
using System.Text;
using OpenMcdf;

namespace DocxportNet.Doc;

/// <summary>Writes one unformatted Unicode main story into a Word binary document.</summary>
internal static class DocPlainTextWriter
{
    private const int PageSize = 512;

    public static void Write(Stream output, string text)
    {
        if (output == null) throw new ArgumentNullException(nameof(output));
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (!output.CanWrite) throw new ArgumentException("The output stream must be writable.", nameof(output));
        if (!text.EndsWith("\r", StringComparison.Ordinal)) text += '\r';

        var encoded = Encoding.Unicode.GetBytes(text);
        const int textOffset = 2048;
        var fcMac = checked(textOffset + encoded.Length);
        var chpxPage = Align(fcMac);
        var paragraphEnds = new List<int>();
        for (var i = 0; i < text.Length; i++)
            if (text[i] == '\r') paragraphEnds.Add(checked(textOffset + (i + 1) * 2));
        var paragraphPages = (paragraphEnds.Count + 28) / 29;
        var word = new byte[checked(chpxPage + PageSize * (1 + paragraphPages))];
        encoded.CopyTo(word, textOffset);

        // One unformatted character run.
        U32(word, chpxPage, textOffset);
        U32(word, chpxPage + 4, fcMac);
        word[chpxPage + 511] = 1;

        // One paragraph boundary per paragraph. BxPap.bOffset=0 uses default style.
        var previous = textOffset;
        for (var page = 0; page < paragraphPages; page++)
        {
            var start = page * 29;
            var count = Math.Min(29, paragraphEnds.Count - start);
            var offset = chpxPage + PageSize * (page + 1);
            U32(word, offset, previous);
            for (var i = 0; i < count; i++)
            {
                previous = paragraphEnds[start + i];
                U32(word, offset + (i + 1) * 4, previous);
            }
            word[offset + 511] = checked((byte)count);
        }

        var fib = new DocFibWriter(word, textOffset, fcMac, text.Length);
        using var table = new MemoryStream();
        fib.AddTableBlock(table, 1, DocDefaultStructures.CreateStyleSheet());
        fib.AddTableBlock(table, 6, DocDefaultStructures.CreateSectionTable(text.Length));

        var chpx = new byte[12];
        U32(chpx, 0, textOffset);
        U32(chpx, 4, fcMac);
        U32(chpx, 8, chpxPage / PageSize);
        fib.AddTableBlock(table, 12, chpx);

        var papx = new byte[(paragraphPages * 2 + 1) * 4];
        U32(papx, 0, textOffset);
        for (var page = 0; page < paragraphPages; page++)
        {
            U32(papx, (page + 1) * 4,
                paragraphEnds[Math.Min((page + 1) * 29, paragraphEnds.Count) - 1]);
            U32(papx, (paragraphPages + 1 + page) * 4, chpxPage / PageSize + 1 + page);
        }
        fib.AddTableBlock(table, 13, papx);

        var clx = new byte[21];
        clx[0] = 2;
        U32(clx, 1, 16);
        U32(clx, 9, text.Length);
        U32(clx, 15, textOffset);
        fib.AddTableBlock(table, 33, clx);

        using var storage = RootStorage.Create(output, OpenMcdf.Version.V3, StorageModeFlags.LeaveOpen);
        WriteStream(storage, "WordDocument", word);
        WriteStream(storage, "1Table", table.ToArray());
    }

    private static void WriteStream(RootStorage storage, string name, byte[] bytes)
    {
        using var stream = storage.CreateStream(name);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static int Align(int value) => checked((value + PageSize - 1) / PageSize * PageSize);
    private static void U32(byte[] target, int offset, int value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(target.AsSpan(offset), checked((uint)value));
}
