using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes STSH headers and length-prefixed styles without parsing style bodies.</summary>
internal static class DocStyleSheetNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var start = location.Offset.Value;
        var end = checked(start + location.Length.Value);
        using var stream = structure.OpenStream(location.StreamName);
        if (start < 0 || end > stream.Length || location.Length < 20)
            throw new InvalidDataException("The stylesheet range is invalid.");
        var stsh = new DocStructureNode("STSH", "StyleSheet", location.StreamName, start, location.Length);
        stream.Position = start;
        var headerLength = ReadU16(stream);
        if (headerLength < 18 || 2L + headerLength > location.Length)
            throw new InvalidDataException("The stylesheet header is invalid.");
        var header = new DocStructureNode("LPStshi", "StyleSheetHeader", location.StreamName,
            start, 2L + headerLength);
        header.Attributes["headerBytes"] = headerLength.ToString(CultureInfo.InvariantCulture);
        var fields = structure.ReadRange(location.StreamName, start, 2 + headerLength);
        var parsedHeader = DocStyleSheetHeader.Read(fields);
        var styleCount = parsedHeader.StyleCount;
        var baseSize = parsedHeader.BaseSize;
        if (styleCount < 15 || styleCount >= 0x0FFE)
            throw new InvalidDataException("The stylesheet style count is invalid.");
        var stshi = new DocStructureNode("STSHI", "StyleSheetInformation", location.StreamName,
            start + 2, headerLength);
        var stshif = new DocStructureNode("Stshif", "StyleSheetFields", location.StreamName,
            start + 2, 18);
        stshif.Attributes["styleCount"] = styleCount.ToString(CultureInfo.InvariantCulture);
        stshif.Attributes["stdBaseBytes"] = baseSize.ToString(CultureInfo.InvariantCulture);
        stshi.Children.Add(stshif);
        header.Children.Add(stshi);
        stsh.Children.Add(header);
        stsh.Attributes["styleCount"] = styleCount.ToString(CultureInfo.InvariantCulture);

        var cursor = start + 2L + headerLength;
        for (var i = 0; i < styleCount; i++)
        {
            if (cursor > end - 2)
                throw new InvalidDataException("A stylesheet entry extends beyond the stylesheet.");
            stream.Position = cursor;
            var styleBytes = ReadU16(stream);
            if (styleBytes > 0x7FFF || styleBytes > end - cursor - 2)
                throw new InvalidDataException("A style definition has an invalid size.");
            var total = 2L + styleBytes + (styleBytes & 1);
            if (total > end - cursor)
                throw new InvalidDataException("A style definition's padding exceeds the stylesheet.");
            var style = new DocStructureNode("LPStd", $"Style{i}", location.StreamName, cursor, total);
            style.Attributes["styleIndex"] = i.ToString(CultureInfo.InvariantCulture);
            style.Attributes["styleBytes"] = styleBytes.ToString(CultureInfo.InvariantCulture);
            if (styleBytes != 0)
                style.Children.Add(new DocStructureNode("STD", "StyleDefinition", location.StreamName,
                    cursor + 2, styleBytes));
            stsh.Children.Add(style);
            cursor += total;
        }
        location.Children.Add(stsh);
    }

    private static ushort ReadU16(Stream stream)
    {
        var bytes = new byte[2];
        if (stream.Read(bytes, 0, 2) != 2)
            throw new InvalidDataException("The stylesheet ended unexpectedly.");
        return BinaryPrimitives.ReadUInt16LittleEndian(bytes);
    }
}
