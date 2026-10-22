using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Text;
using DocxportNet.Visitors.Pdf;
using PdfSharp.Internal.OpenType;

/// <summary>Indexes the OpenType names of font files in CLI-supplied directories.</summary>
internal sealed class DxpDirectoryPdfFontProvider : IDxpPdfFontProvider
{
    private readonly Dictionary<string, List<Face>> _families = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentDictionary<string, Lazy<OpenTypeFontDescriptor?>> _descriptors = new();
    private readonly ConcurrentDictionary<string, byte[]> _fontBytes = new();

    public DxpDirectoryPdfFontProvider(IEnumerable<string> directories)
    {
        foreach (string directory in directories)
        {
            foreach (string path in Directory.EnumerateFiles(directory, "*", new EnumerationOptions
                     { RecurseSubdirectories = true, IgnoreInaccessible = true })
                     .OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            {
                if (!IsSupportedExtension(path))
                    continue;
                try
                {
                    byte[] bytes = File.ReadAllBytes(path);
                    if (!TryReadNames(bytes, out string? family, out string? style))
                        continue;
                    bool bold = style.Contains("Bold", StringComparison.OrdinalIgnoreCase);
                    bool italic = style.Contains("Italic", StringComparison.OrdinalIgnoreCase) ||
                                  style.Contains("Oblique", StringComparison.OrdinalIgnoreCase);
                    if (!_families.TryGetValue(family, out var faces))
                        _families[family] = faces = [];
                    faces.Add(new Face(family, path, bold, italic));
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
                .Select(pair => BestFace(pair.Value.Where(f => HasGlyph(f, character)), bold, italic))
                .FirstOrDefault(candidate => candidate != null);
        }
        return face == null ? null : new DxpPdfFontMatch(face.Family, Materialize(face, bold, italic));
    }

    private DxpPdfFontFace Materialize(Face face, bool bold, bool italic) =>
        new(face.Path, _fontBytes.GetOrAdd(face.Path, File.ReadAllBytes),
            bold && !face.Bold, italic && !face.Italic);

    private bool HasGlyph(Face face, Rune character)
    {
        var descriptor = _descriptors.GetOrAdd(face.Path, path => new Lazy<OpenTypeFontDescriptor?>(() =>
        {
            try
            {
                var source = OpenTypeFontSource.GetOrCreateFrom(_fontBytes.GetOrAdd(path, File.ReadAllBytes));
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
        Path.GetExtension(path).Equals(".otf", StringComparison.OrdinalIgnoreCase);

    private static bool TryReadNames(byte[] bytes, out string family, out string style)
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
        family = names.GetValueOrDefault(16).Value ?? names.GetValueOrDefault(1).Value ?? "";
        style = names.GetValueOrDefault(17).Value ?? names.GetValueOrDefault(2).Value ?? "Regular";
        return family.Length > 0;
    }

    private static bool CanRead(byte[] bytes, int offset, int length) =>
        offset >= 0 && length >= 0 && offset <= bytes.Length && length <= bytes.Length - offset;

    private static ushort U16(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(offset, 2));

    private static uint U32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(offset, 4));

    private sealed record Face(string Family, string Path, bool Bold, bool Italic);
}
