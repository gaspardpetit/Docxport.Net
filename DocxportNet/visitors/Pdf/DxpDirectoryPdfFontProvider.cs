using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;
using PdfSharp.Internal.OpenType;

namespace DocxportNet.Visitors.Pdf;

/// <summary>Indexes the OpenType names and glyph coverage of font files in directories.</summary>
public sealed class DxpDirectoryPdfFontProvider : IDxpPdfFontProvider
{
    private readonly Dictionary<string, List<Face>> _families = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, Lazy<OpenTypeFontDescriptor?>> _descriptors = new();
    private readonly ConcurrentDictionary<string, byte[]> _fontBytes = new();

    public DxpDirectoryPdfFontProvider(IEnumerable<string> directories)
    {
        foreach (string directory in directories)
        {
            if (!Directory.Exists(directory)) continue;
            foreach (string path in Directory.EnumerateFiles(directory, "*", new EnumerationOptions
                     { RecurseSubdirectories = true, IgnoreInaccessible = true })
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                if (!IsSupportedExtension(path))
                    continue;
                try
                {
                    foreach (var (index, family, style) in ReadFaceNames(path))
                    {
                        bool bold = style.Contains("Bold", StringComparison.OrdinalIgnoreCase);
                        bool italic = style.Contains("Italic", StringComparison.OrdinalIgnoreCase) ||
                                      style.Contains("Oblique", StringComparison.OrdinalIgnoreCase);
                        if (!_families.TryGetValue(family, out var faces))
                            _families[family] = faces = [];
                        faces.Add(new Face(family, path, index, bold, italic));
                    }
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    public int FaceCount => _families.Values.Sum(faces => faces.Count);
    public IReadOnlyCollection<string> Families => _families.Keys;

    public bool IsAvailable(string familyName, bool bold, bool italic) =>
        _families.TryGetValue(familyName, out var faces) && BestFace(faces, bold, italic) != null;

    public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic)
    {
        if (!_families.TryGetValue(familyName, out var faces)) return null;
        Face? face = BestFace(faces, bold, italic);
        return face == null ? null : Materialize(face, bold, italic);
    }

    public DxpPdfFontMatch? ResolveForCharacter(string familyName, bool bold, bool italic, Rune character)
    {
        // Do not replace unrelated system fonts just because an extra directory is present.
        if (!_families.TryGetValue(familyName, out var requested)) return null;
        Face? face = BestFace(requested.Where(f => HasGlyph(f, character)), bold, italic);
        if (face == null)
        {
            face = _families.Where(pair => !pair.Key.Equals(familyName, StringComparison.OrdinalIgnoreCase))
                .OrderBy(pair => FallbackRank(pair.Key, character))
                .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
                .Select(pair => BestFace(pair.Value.Where(f => HasGlyph(f, character)), bold, italic))
                .FirstOrDefault(candidate => candidate != null);
        }
        return face == null ? null : new DxpPdfFontMatch(face.Family, Materialize(face, bold, italic));
    }

    /// <summary>Finds a covering installed face when the requested family is absent.</summary>
    public DxpPdfFontMatch? ResolveAnyForCharacter(bool bold, bool italic, Rune character)
    {
        Face? face = _families.OrderBy(pair => FallbackRank(pair.Key, character))
            .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .Select(pair => BestFace(pair.Value.Where(f => HasGlyph(f, character)), bold, italic))
            .FirstOrDefault(candidate => candidate != null);
        return face == null ? null : new DxpPdfFontMatch(face.Family, Materialize(face, bold, italic));
    }

    private static int FallbackRank(string family, Rune character)
    {
        // Keep fallbacks stable across directory ordering and prefer the system's
        // broad CJK faces to unrelated decorative fonts that happen to cover Han.
        if (character.Value is >= 0x3400 and <= 0x9FFF or >= 0x20000 and <= 0x3134F)
        {
            string[] preferred = ["Microsoft JhengHei", "Microsoft YaHei", "Noto Sans CJK SC",
                "Noto Sans CJK TC", "Noto Sans CJK JP", "Noto Sans JP", "Yu Gothic"];
            int rank = Array.FindIndex(preferred, name => name.Equals(family, StringComparison.OrdinalIgnoreCase));
            return rank >= 0 ? rank : preferred.Length;
        }
        return 0;
    }

    private DxpPdfFontFace Materialize(Face face, bool bold, bool italic) =>
        new(FaceKey(face), FontBytes(face),
            bold && !face.Bold, italic && !face.Italic);

    private static string FaceKey(Face face) => face.Path + "#" + face.Index;

    private byte[] FontBytes(Face face) => _fontBytes.GetOrAdd(FaceKey(face), _ =>
        face.Index < 0 ? File.ReadAllBytes(face.Path) :
            ExtractCollectionFace(File.ReadAllBytes(face.Path), face.Index));

    private bool HasGlyph(Face face, Rune character)
    {
        var descriptor = _descriptors.GetOrAdd(FaceKey(face), _ => new Lazy<OpenTypeFontDescriptor?>(() =>
        {
            try
            {
                var source = OpenTypeFontSource.GetOrCreateFrom(FontBytes(face));
                return OpenTypeFontFace.GetOrCreateFrom(source).OTDescriptor;
            }
            catch (InvalidOperationException) { return null; }
            catch (InvalidDataException) { return null; }
            catch (IOException) { return null; }
        })).Value;
        return descriptor?.CodepointToGlyphIndex(character.Value) > 0;
    }

    private static Face? BestFace(IEnumerable<Face> faces, bool bold, bool italic)
    {
        // A bold-only face must not satisfy a regular request. Missing requested
        // styles may be synthesized when the directory has no exact face.
        return faces.Where(f => (!f.Bold || bold) && (!f.Italic || italic))
            .OrderByDescending(f => (f.Bold == bold ? 1 : 0) + (f.Italic == italic ? 1 : 0))
            .FirstOrDefault();
    }

    private static bool IsSupportedExtension(string path) =>
        Path.GetExtension(path).Equals(".ttf", StringComparison.OrdinalIgnoreCase) ||
        Path.GetExtension(path).Equals(".otf", StringComparison.OrdinalIgnoreCase) ||
        Path.GetExtension(path).Equals(".ttc", StringComparison.OrdinalIgnoreCase);

    private static IEnumerable<(int Index, string Family, string Style)> ReadFaceNames(string path)
    {
        using var stream = File.OpenRead(path);
        byte[] header = new byte[12];
        if (stream.Read(header, 0, header.Length) != header.Length) yield break;
        if (U32(header, 0) != 0x74746366) // 'ttcf'
        {
            var names = ReadNamesAt(stream, 0);
            if (names is { } result) yield return (-1, result.Family, result.Style);
            yield break;
        }
        uint count = U32(header, 8);
        if (count is 0 or > 256 || stream.Length < 12 + count * 4) yield break;
        byte[] offsets = new byte[count * 4];
        stream.ReadExactly(offsets);
        for (int index = 0; index < count; index++)
        {
            uint offset = U32(offsets, index * 4);
            if (offset > stream.Length - 12) continue;
            var names = ReadNamesAt(stream, offset);
            if (names is { } result) yield return (index, result.Family, result.Style);
        }
    }

    private static (string Family, string Style)? ReadNamesAt(FileStream stream, long faceOffset)
    {
        stream.Position = faceOffset;
        byte[] header = new byte[12];
        if (stream.Read(header, 0, header.Length) != header.Length) return null;
        int count = U16(header, 4);
        if (count is 0 or > 4096 || stream.Length - stream.Position < count * 16) return null;
        byte[] records = new byte[count * 16];
        stream.ReadExactly(records);
        for (int index = 0; index < count; index++)
        {
            int record = index * 16;
            if (U32(records, record) != 0x6E616D65) continue; // 'name'
            uint offset = U32(records, record + 8);
            uint length = U32(records, record + 12);
            if (length is 0 or > 4_000_000 || offset > stream.Length || length > stream.Length - offset)
                return null;
            byte[] font = new byte[28 + length];
            BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(0, 4), 0x00010000);
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(4, 2), 1);
            BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(12, 4), 0x6E616D65);
            BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(20, 4), 28);
            BinaryPrimitives.WriteUInt32BigEndian(font.AsSpan(24, 4), length);
            stream.Position = offset;
            stream.ReadExactly(font.AsSpan(28));
            return TryReadNames(font, out string family, out string style) ? (family, style) : null;
        }
        return null;
    }

    private static byte[] ExtractCollectionFace(byte[] collection, int faceIndex)
    {
        if (!CanRead(collection, 0, 12) || U32(collection, 0) != 0x74746366)
            throw new InvalidDataException("Expected a TrueType collection.");
        uint faceCount = U32(collection, 8);
        if (faceIndex < 0 || faceIndex >= faceCount || faceCount > 256 ||
            !CanRead(collection, 12, checked((int)faceCount * 4)))
            throw new InvalidDataException("Invalid TrueType collection face index.");
        int fontOffset = checked((int)U32(collection, 12 + faceIndex * 4));
        if (!CanRead(collection, fontOffset, 12))
            throw new InvalidDataException("Invalid TrueType collection face offset.");
        int count = U16(collection, fontOffset + 4);
        if (!CanRead(collection, fontOffset + 12, count * 16))
            throw new InvalidDataException("Invalid TrueType collection table directory.");
        var tables = new List<(byte[] Tag, byte[] Data)>();
        for (int index = 0; index < count; index++)
        {
            int record = fontOffset + 12 + index * 16;
            uint offset = U32(collection, record + 8);
            uint length = U32(collection, record + 12);
            if (offset > int.MaxValue || length > int.MaxValue ||
                !CanRead(collection, (int)offset, (int)length))
                throw new InvalidDataException("Invalid TrueType collection table offset.");
            byte[] data = collection.AsSpan((int)offset, (int)length).ToArray();
            byte[] tag = collection.AsSpan(record, 4).ToArray();
            if (tag.AsSpan().SequenceEqual("head"u8) && data.Length >= 12)
                data.AsSpan(8, 4).Clear();
            tables.Add((tag, data));
        }
        int lengthBytes = checked(12 + count * 16 + tables.Sum(table => (table.Data.Length + 3) & ~3));
        var face = new byte[lengthBytes];
        collection.AsSpan(fontOffset, 4).CopyTo(face);
        BinaryPrimitives.WriteUInt16BigEndian(face.AsSpan(4, 2), (ushort)count);
        int power = 1, selector = 0;
        while (power * 2 <= count) { power *= 2; selector++; }
        BinaryPrimitives.WriteUInt16BigEndian(face.AsSpan(6, 2), (ushort)(power * 16));
        BinaryPrimitives.WriteUInt16BigEndian(face.AsSpan(8, 2), (ushort)selector);
        BinaryPrimitives.WriteUInt16BigEndian(face.AsSpan(10, 2), (ushort)(count * 16 - power * 16));
        int destination = 12 + count * 16, headOffset = -1;
        for (int index = 0; index < tables.Count; index++)
        {
            var (tag, data) = tables[index];
            var record = face.AsSpan(12 + index * 16, 16);
            tag.CopyTo(record);
            BinaryPrimitives.WriteUInt32BigEndian(record.Slice(4, 4), Checksum(data));
            BinaryPrimitives.WriteUInt32BigEndian(record.Slice(8, 4), (uint)destination);
            BinaryPrimitives.WriteUInt32BigEndian(record.Slice(12, 4), (uint)data.Length);
            data.CopyTo(face.AsSpan(destination));
            if (tag.AsSpan().SequenceEqual("head"u8)) headOffset = destination;
            destination += (data.Length + 3) & ~3;
        }
        if (headOffset < 0) throw new InvalidDataException("TrueType collection has no head table.");
        BinaryPrimitives.WriteUInt32BigEndian(face.AsSpan(headOffset + 8, 4),
            unchecked(0xB1B0AFBAu - Checksum(face)));
        return face;
    }

    private static uint Checksum(ReadOnlySpan<byte> data)
    {
        uint sum = 0;
        for (int index = 0; index < data.Length; index += 4)
        {
            uint word = 0;
            for (int part = 0; part < 4 && index + part < data.Length; part++)
                word |= (uint)data[index + part] << (24 - part * 8);
            sum = unchecked(sum + word);
        }
        return sum;
    }

    internal static bool TryReadNames(byte[] bytes, out string family, out string style)
    {
        family = style = "";
        if (!CanRead(bytes, 0, 12))
            return false;
        uint signature = U32(bytes, 0);
        if (signature is not (0x00010000 or 0x4F54544F or 0x74727565))
            return false; // TrueType/OpenType single faces; collections need face selection.
        int tableCount = U16(bytes, 4);
        if (!CanRead(bytes, 12, tableCount * 16))
            return false;
        int nameOffset = -1;
        int nameLength = 0;
        for (int i = 0; i < tableCount; i++)
        {
            int record = 12 + i * 16;
            if (U32(bytes, record) != 0x6E616D65) // 'name'
                continue;
            uint offset = U32(bytes, record + 8);
            uint length = U32(bytes, record + 12);
            if (offset > int.MaxValue || length > int.MaxValue || !CanRead(bytes, (int)offset, (int)length))
                return false;
            nameOffset = (int)offset;
            nameLength = (int)length;
            break;
        }
        if (nameOffset < 0 || nameLength < 6)
            return false;
        int count = U16(bytes, nameOffset + 2);
        int strings = U16(bytes, nameOffset + 4);
        if (6 + count * 12 > nameLength)
            return false;
        var names = new Dictionary<int, (int Rank, string Value)>();
        for (int i = 0; i < count; i++)
        {
            int record = nameOffset + 6 + i * 12;
            int platform = U16(bytes, record);
            int language = U16(bytes, record + 4);
            int id = U16(bytes, record + 6);
            if (id is not (1 or 2 or 16 or 17) || platform is not (0 or 1 or 3))
                continue;
            int length = U16(bytes, record + 8);
            int offset = nameOffset + strings + U16(bytes, record + 10);
            if (!CanRead(bytes, offset, length) || offset + length > nameOffset + nameLength)
                continue;
            string value;
            if (platform is 0 or 3)
            {
                if (length % 2 != 0) continue;
                value = Encoding.BigEndianUnicode.GetString(bytes, offset, length);
            }
            else
                value = Encoding.Latin1.GetString(bytes, offset, length);
            value = value.Trim('\0', ' ');
            if (value.Length == 0)
                continue;
            int rank = platform == 3 && language == 0x0409 ? 3 : platform == 0 ? 2 : 1;
            if (!names.TryGetValue(id, out var old) || rank > old.Rank)
                names[id] = (rank, value);
        }
        string? typographicFamily = names.GetValueOrDefault(16).Value;
        string? legacyFamily = names.GetValueOrDefault(1).Value;
        // Some faces encode "Display" or "Narrow" in the typographic style
        // while applications request it as part of the legacy family name.
        family = typographicFamily is { Length: > 0 } && legacyFamily is { Length: > 0 } &&
                 legacyFamily.StartsWith(typographicFamily + " ", StringComparison.OrdinalIgnoreCase)
            ? legacyFamily : typographicFamily ?? legacyFamily ?? "";
        style = names.GetValueOrDefault(17).Value ?? names.GetValueOrDefault(2).Value ?? "Regular";
        return family.Length > 0;
    }

    private static bool CanRead(byte[] bytes, int offset, int length) =>
        offset >= 0 && length >= 0 && offset <= bytes.Length && length <= bytes.Length - offset;

    private static ushort U16(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));

    private static uint U32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4));

    private sealed record Face(string Family, string Path, int Index, bool Bold, bool Italic);
}
