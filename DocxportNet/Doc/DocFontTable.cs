using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace DocxportNet.Doc;

/// <summary>A named font referenced by a DOC font index.</summary>
public sealed record DocFontDefinition(int Index, string Name,
    byte FamilyPitch = 0, short Weight = 400, byte Charset = 0,
    byte[]? Panose = null, byte[]? Signature = null,
    string? AlternateName = null);

internal static class DocFontTable
{
    public static IReadOnlyList<DocFontDefinition> ReadOpenXml(MainDocumentPart? main)
    {
        var xml = main?.FontTablePart?.Fonts?.OuterXml;
        if (string.IsNullOrEmpty(xml)) return [];
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var root = XElement.Parse(xml);
        var result = new List<DocFontDefinition>();
        foreach (var element in root.Elements(w + "font"))
        {
            var name = (string?)element.Attribute(w + "name");
            if (string.IsNullOrWhiteSpace(name)) continue;
            string? Value(string child) => (string?)element.Element(w + child)?.Attribute(w + "val");
            var family = Value("family") switch
            {
                "roman" => 1, "swiss" => 2, "modern" => 3,
                "script" => 4, "decorative" => 5, _ => 0
            };
            var pitch = Value("pitch") switch
            {
                "fixed" => 1, "variable" => 2, _ => 0
            };
            var familyPitch = checked((byte)((family << 4) |
                (element.Element(w + "notTrueType") == null ? 4 : 0) | pitch));
            var charset = byte.TryParse(Value("charset"), NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out var parsedCharset) ? parsedCharset : (byte)0;
            static byte[]? Hex(string? value, int length) =>
                value?.Length == length * 2 &&
                value.All(Uri.IsHexDigit) ? DocBinaryCompat.Hex(value) : null;
            var panose = Hex(Value("panose1"), 10);
            var signatureElement = element.Element(w + "sig");
            byte[]? signature = null;
            if (signatureElement != null)
            {
                var fields = new[] { "usb0", "usb1", "usb2", "usb3", "csb0", "csb1" };
                var values = fields.Select(x => (string?)signatureElement.Attribute(w + x)).ToArray();
                if (values.All(x => x?.Length == 8 && x.All(Uri.IsHexDigit)))
                {
                    signature = new byte[24];
                    for (var i = 0; i < fields.Length; i++)
                        BinaryPrimitives.WriteUInt32LittleEndian(signature.AsSpan(i * 4),
                            uint.Parse(values[i]!, NumberStyles.HexNumber,
                                CultureInfo.InvariantCulture));
                }
            }
            result.Add(new DocFontDefinition(result.Count, name, familyPitch, 400,
                charset, panose, signature, Value("altName")));
        }
        return result;
    }

    public static IReadOnlyList<DocFontDefinition> Read(DocStructure structure)
    {
        var location = structure.FindLocation("FontTable");
        if (location?.IsPresent != true) return Array.Empty<DocFontDefinition>();
        var bytes = structure.ReadRange(location.StreamName, location.Offset,
            checked((int)location.Length));
        if (bytes.Length < 4)
            throw new InvalidDataException("The DOC font table is truncated.");
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        if (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2)) != 0)
            throw new InvalidDataException("The DOC font table has extra entry data.");
        var result = new List<DocFontDefinition>(count);
        var cursor = 4;
        for (var i = 0; i < count; i++)
        {
            if (cursor >= bytes.Length)
                throw new InvalidDataException("A DOC font record is truncated.");
            var length = bytes[cursor++];
            if (length < 41 || length > bytes.Length - cursor)
                throw new InvalidDataException("A DOC font record has an invalid length.");
            var nameStart = cursor + 39;
            var nameEnd = cursor + length;
            var terminator = nameStart;
            while (terminator + 1 < nameEnd &&
                BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(terminator)) != 0)
                terminator += 2;
            if (terminator + 1 >= nameEnd)
                throw new InvalidDataException("A DOC font name has no terminator.");
            var alternateIndex = bytes[cursor + 4];
            string? alternateName = null;
            if (alternateIndex != 0)
            {
                var alternateStart = checked(nameStart + alternateIndex * 2);
                if (alternateStart <= terminator || alternateStart >= nameEnd - 1)
                    throw new InvalidDataException("A DOC alternate font name has an invalid offset.");
                var alternateEnd = alternateStart;
                while (alternateEnd + 1 < nameEnd &&
                    BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(alternateEnd)) != 0)
                    alternateEnd += 2;
                if (alternateEnd + 1 >= nameEnd)
                    throw new InvalidDataException("A DOC alternate font name has no terminator.");
                alternateName = Encoding.Unicode.GetString(bytes,
                    alternateStart, alternateEnd - alternateStart);
            }
            result.Add(new DocFontDefinition(i,
                Encoding.Unicode.GetString(bytes, nameStart, terminator - nameStart),
                bytes[cursor], BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(cursor + 1)),
                bytes[cursor + 3], bytes.AsSpan(cursor + 5, 10).ToArray(),
                bytes.AsSpan(cursor + 15, 24).ToArray(), alternateName));
            cursor += length;
        }
        return result;
    }

    public static byte[] Write(IReadOnlyList<string> names,
        IReadOnlyList<DocFontDefinition>? metadata = null)
    {
        if (names.Count > 0x7FF0)
            throw new InvalidDataException("The DOC font table has too many entries.");
        using var output = new MemoryStream();
        output.WriteByte((byte)names.Count);
        output.WriteByte((byte)(names.Count >> 8));
        output.WriteByte(0); output.WriteByte(0);
        var byName = metadata?.GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
        foreach (var name in names)
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new InvalidDataException("A DOC font name is empty.");
            DocFontDefinition? font = null;
            if (byName != null) byName.TryGetValue(name, out font);
            var alternate = font?.AlternateName;
            var nameBytes = Encoding.Unicode.GetBytes(name + "\0" +
                (alternate == null ? "" : alternate + "\0"));
            var length = checked(39 + nameBytes.Length);
            if (length > 255)
                throw new InvalidDataException("A DOC font name is too long.");
            output.WriteByte(checked((byte)length));
            var record = new byte[length];
            record[0] = font?.FamilyPitch ?? 0;
            BinaryPrimitives.WriteInt16LittleEndian(record.AsSpan(1), font?.Weight ?? 400);
            record[3] = font?.Charset ?? 0;
            if (alternate != null)
                record[4] = checked((byte)(name.Length + 1));
            if (font?.Panose is { Length: 10 } panose)
                panose.CopyTo(record, 5);
            if (font?.Signature is { Length: 24 } signature)
                signature.CopyTo(record, 15);
            nameBytes.CopyTo(record, 39);
            output.Write(record, 0, record.Length);
        }
        return output.ToArray();
    }
}
