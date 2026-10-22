using System.Buffers.Binary;
using System.Text;
using DocxportNet.Visitors.Pdf;

/// <summary>Indexes the OpenType names of font files in CLI-supplied directories.</summary>
internal sealed class DxpDirectoryPdfFontProvider : IDxpPdfFontProvider
{
    private readonly Dictionary<string, List<Face>> _families = new(StringComparer.OrdinalIgnoreCase);

    public DxpDirectoryPdfFontProvider(IEnumerable<string> directories)
    {
        foreach (string directory in directories)
        {
            foreach (string path in Directory.EnumerateFiles(directory, "*", new EnumerationOptions
                     { RecurseSubdirectories = true, IgnoreInaccessible = true }))
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
                    faces.Add(new Face(path, bold, italic));
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }
    }

    public int FaceCount => _families.Values.Sum(faces => faces.Count);
    public IReadOnlyCollection<string> Families => _families.Keys;

    public DxpPdfFontFace? Resolve(string familyName, bool bold, bool italic)
    {
        if (!_families.TryGetValue(familyName, out var faces))
            return null;
        // A bold-only face must not satisfy a regular request. Missing requested
        // styles may be synthesized when the directory has no exact face.
        Face? face = faces.Where(f => (!f.Bold || bold) && (!f.Italic || italic))
            .OrderByDescending(f => (f.Bold == bold ? 1 : 0) + (f.Italic == italic ? 1 : 0))
            .FirstOrDefault();
        if (face is null)
            return null;
        return new DxpPdfFontFace(face.Path, File.ReadAllBytes(face.Path),
            bold && !face.Bold, italic && !face.Italic);
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

    private sealed record Face(string Path, bool Bold, bool Italic);
}
