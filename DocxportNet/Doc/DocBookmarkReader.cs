using System.Buffers.Binary;
using System.Text;

namespace DocxportNet.Doc;

/// <summary>A standard DOC bookmark range in global character positions.</summary>
public sealed record DocBookmark(string Name, uint CpStart, uint CpEnd);

internal static class DocBookmarkReader
{
    public static IReadOnlyList<DocBookmark> Read(DocTextIndex index)
    {
        var structure = index.Structure;
        DocLocation? Location(string name) => structure.Locations.FirstOrDefault(x =>
            x.Name == name && x.IsPresent);
        var namesLocation = Location("BookmarkNames");
        var startsLocation = Location("BookmarkStarts");
        var endsLocation = Location("BookmarkEnds");
        if (namesLocation == null && startsLocation == null && endsLocation == null)
            return Array.Empty<DocBookmark>();
        if (namesLocation == null || startsLocation == null || endsLocation == null)
            throw new InvalidDataException("The DOC bookmark tables are incomplete.");
        byte[] Bytes(DocLocation location) => structure.ReadRange(location.StreamName,
            location.Offset, checked((int)location.Length));
        var names = ReadNames(Bytes(namesLocation));
        var starts = Bytes(startsLocation);
        var ends = Bytes(endsLocation);
        var count = names.Count;
        if (starts.Length != checked(count * 8 + 4) ||
            ends.Length != checked((count + 1) * 4))
            throw new InvalidDataException("The DOC bookmark tables have inconsistent counts.");
        var result = new DocBookmark[count];
        for (var i = 0; i < count; i++)
        {
            var start = BinaryPrimitives.ReadUInt32LittleEndian(starts.AsSpan(i * 4));
            var endIndex = BinaryPrimitives.ReadUInt16LittleEndian(
                starts.AsSpan((count + 1) * 4 + i * 4));
            if (endIndex >= count)
                throw new InvalidDataException("A DOC bookmark has an invalid end index.");
            var end = BinaryPrimitives.ReadUInt32LittleEndian(ends.AsSpan(endIndex * 4));
            if (end < start)
                throw new InvalidDataException("A DOC bookmark ends before it starts.");
            result[i] = new DocBookmark(names[i], start, end);
        }
        return result;
    }

    private static IReadOnlyList<string> ReadNames(byte[] bytes)
    {
        if (bytes.Length < 6 || BinaryPrimitives.ReadUInt16LittleEndian(bytes) != 0xFFFF)
            throw new NotSupportedException("Legacy ANSI DOC bookmark names are unsupported.");
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2));
        var extra = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(4));
        var result = new string[count];
        var offset = 6;
        for (var i = 0; i < count; i++)
        {
            if (offset + 2 > bytes.Length)
                throw new InvalidDataException("A DOC bookmark name is truncated.");
            var length = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset));
            offset += 2;
            var size = checked(length * 2);
            if (offset + size + extra > bytes.Length)
                throw new InvalidDataException("A DOC bookmark name is truncated.");
            result[i] = Encoding.Unicode.GetString(bytes, offset, size);
            offset += size + extra;
        }
        if (offset != bytes.Length)
            throw new InvalidDataException("The DOC bookmark name table has trailing data.");
        return result;
    }
}
