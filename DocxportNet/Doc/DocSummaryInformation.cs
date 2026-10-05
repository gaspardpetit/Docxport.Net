using System.Buffers.Binary;
using System.Text;

namespace DocxportNet.Doc;

internal static class DocSummaryInformation
{
    private const string StreamName = "\u0005SummaryInformation";

    public static byte[] Create(string? title, string? subject, string? author,
        string? keywords, string? comments, string? lastAuthor,
        int? pageCount = null, int? wordCount = null, int? characterCount = null,
        string? revisionNumber = null)
    {
        var properties = new List<(uint Id, byte[] Value)>
        {
            (1, new byte[] { 2, 0, 0, 0, 0xE9, 0xFD, 0, 0 })
        };
        Add(2, title);
        Add(3, subject);
        Add(4, author);
        Add(5, keywords);
        Add(6, comments);
        Add(8, lastAuthor);
        Add(9, revisionNumber);
        AddInteger(14, pageCount);
        AddInteger(15, wordCount);
        AddInteger(16, characterCount);
        var directoryLength = 8 + properties.Count * 8;
        var bytes = new byte[48 + directoryLength + properties.Sum(x => x.Value.Length)];
        BinaryPrimitives.WriteUInt16LittleEndian(bytes, 0xFFFE);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(24), 1);
        new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9")
            .ToByteArray().CopyTo(bytes, 28);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(44), 48);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(48),
            checked((uint)(bytes.Length - 48)));
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(52),
            checked((uint)properties.Count));
        var valueOffset = directoryLength;
        for (var i = 0; i < properties.Count; i++)
        {
            var entry = bytes.AsSpan(56 + i * 8);
            BinaryPrimitives.WriteUInt32LittleEndian(entry, properties[i].Id);
            BinaryPrimitives.WriteUInt32LittleEndian(entry.Slice(4),
                checked((uint)valueOffset));
            properties[i].Value.CopyTo(bytes, 48 + valueOffset);
            valueOffset += properties[i].Value.Length;
        }
        return bytes;

        void Add(uint id, string? value)
        {
            if (value == null) return;
            var characters = Encoding.UTF8.GetBytes(value + "\0");
            var encoded = new byte[(8 + characters.Length + 3) & ~3];
            BinaryPrimitives.WriteUInt16LittleEndian(encoded, 0x1E);
            BinaryPrimitives.WriteUInt32LittleEndian(encoded.AsSpan(4),
                checked((uint)characters.Length));
            characters.CopyTo(encoded, 8);
            properties.Add((id, encoded));
        }

        void AddInteger(uint id, int? value)
        {
            if (value is not >= 0) return;
            var encoded = new byte[8];
            BinaryPrimitives.WriteUInt16LittleEndian(encoded, 3);
            BinaryPrimitives.WriteInt32LittleEndian(encoded.AsSpan(4), value.Value);
            properties.Add((id, encoded));
        }
    }

    public static string? ReadTitle(DocStructure structure) => ReadProperty(structure, 2);
    public static string? ReadSubject(DocStructure structure) => ReadProperty(structure, 3);
    public static string? ReadAuthor(DocStructure structure) => ReadProperty(structure, 4);
    public static string? ReadKeywords(DocStructure structure) => ReadProperty(structure, 5);
    public static string? ReadComments(DocStructure structure) => ReadProperty(structure, 6);
    public static string? ReadLastAuthor(DocStructure structure) => ReadProperty(structure, 8);
    public static string? ReadRevisionNumber(DocStructure structure) => ReadProperty(structure, 9);
    public static int? ReadPageCount(DocStructure structure) => ReadIntegerProperty(structure, 14);
    public static int? ReadWordCount(DocStructure structure) => ReadIntegerProperty(structure, 15);
    public static int? ReadCharacterCount(DocStructure structure) => ReadIntegerProperty(structure, 16);

    private static int? ReadIntegerProperty(DocStructure structure, uint propertyId)
    {
        var property = FindProperty(structure, propertyId);
        if (property.Length < 8 ||
            BinaryPrimitives.ReadUInt16LittleEndian(property) != 3)
            return null;
        var value = BinaryPrimitives.ReadInt32LittleEndian(property.Slice(4));
        return value >= 0 ? value : null;
    }

    private static string? ReadProperty(DocStructure structure, uint propertyId)
    {
        var property = FindProperty(structure, propertyId, out var codePage);
        if (property.Length < 8 || BinaryPrimitives.ReadUInt16LittleEndian(property) != 0x1E)
            return null;
        var length = BinaryPrimitives.ReadUInt32LittleEndian(property.Slice(4));
        if (length > property.Length - 8) return null;
        var characters = property.Slice(8, (int)length);
        if (codePage == 1200)
        {
            if ((characters.Length & 1) != 0) return null;
            return Encoding.Unicode.GetString(characters.ToArray()).TrimEnd('\0');
        }
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding((int)codePage).GetString(characters.ToArray()).TrimEnd('\0');
    }

    private static ReadOnlySpan<byte> FindProperty(DocStructure structure, uint propertyId)
        => FindProperty(structure, propertyId, out _);

    private static ReadOnlySpan<byte> FindProperty(DocStructure structure, uint propertyId,
        out uint codePage)
    {
        codePage = 1252;
        if (!structure.Root.Children.Any(x => x.Kind == "Stream" && x.Name == StreamName))
            return [];
        using var stream = structure.OpenStream(StreamName);
        if (stream.Length < 80 || stream.Length > int.MaxValue)
            throw new InvalidDataException("The DOC summary information is truncated or too large.");
        var data = new byte[(int)stream.Length];
        stream.ReadExactly(data);
        var span = data.AsSpan();
        if (BinaryPrimitives.ReadUInt16LittleEndian(span) != 0xFFFE ||
            BinaryPrimitives.ReadUInt32LittleEndian(span.Slice(24)) != 1)
            return [];
        var offset = BinaryPrimitives.ReadUInt32LittleEndian(span.Slice(44));
        if (offset > data.Length - 8) return [];
        var set = span.Slice((int)offset);
        var size = BinaryPrimitives.ReadUInt32LittleEndian(set);
        var count = BinaryPrimitives.ReadUInt32LittleEndian(set.Slice(4));
        if (size < 8 || size > set.Length || count > (size - 8) / 8) return [];
        uint propertyOffset = 0;
        for (var i = 0; i < count; i++)
        {
            var entry = set.Slice(8 + (int)i * 8);
            var id = BinaryPrimitives.ReadUInt32LittleEndian(entry);
            var valueOffset = BinaryPrimitives.ReadUInt32LittleEndian(entry.Slice(4));
            if (valueOffset > size - 8) continue;
            if (id == 1 && BinaryPrimitives.ReadUInt16LittleEndian(set.Slice((int)valueOffset)) == 2)
                codePage = BinaryPrimitives.ReadUInt16LittleEndian(set.Slice((int)valueOffset + 4));
            if (id == propertyId) propertyOffset = valueOffset;
        }
        if (propertyOffset == 0 || propertyOffset > size - 8) return [];
        return set.Slice((int)propertyOffset).ToArray();
    }
}
