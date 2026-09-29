using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes section ranges and their section-property byte ranges.</summary>
internal static class DocSectionNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null)
            return;

        var length = location.Length.Value;
        if (length < 20 || (length - 4) % 16 != 0)
            throw new InvalidDataException("The section table has an invalid size.");
        var count = checked((int)((length - 4) / 16));
        var cpOffset = location.Offset.Value;
        var sedOffset = cpOffset + (count + 1L) * 4;
        using var table = structure.OpenStream(location.StreamName);
        using var word = structure.OpenStream("WordDocument");
        if (cpOffset < 0 || length > table.Length - cpOffset)
            throw new InvalidDataException("The section table points outside its stream.");
        if (count == 1)
            _ = DocSectionTable.Read(structure.ReadRange(location.StreamName, cpOffset, checked((int)length)));

        var plc = new DocStructureNode("PlcfSed", "SectionTable", location.StreamName, cpOffset, length);
        plc.Attributes["sectionCount"] = count.ToString(CultureInfo.InvariantCulture);
        table.Position = cpOffset;
        var previousCp = ReadU32(table);
        if (previousCp != 0)
            throw new InvalidDataException("The first section CP must be zero.");

        for (var i = 0; i < count; i++)
        {
            table.Position = cpOffset + (i + 1L) * 4;
            var nextCp = ReadU32(table);
            if (nextCp <= previousCp || nextCp >= 0x80000000)
                throw new InvalidDataException("Section CPs must increase within the valid range.");
            var recordOffset = sedOffset + i * 12L;
            table.Position = recordOffset + 2; // fn is undefined by the format.
            var sepxOffset = unchecked((int)ReadU32(table));
            var sed = new DocStructureNode("Sed", $"Section{i}", location.StreamName, recordOffset, 12);
            sed.Attributes["cpStart"] = previousCp.ToString(CultureInfo.InvariantCulture);
            sed.Attributes["cpEnd"] = nextCp.ToString(CultureInfo.InvariantCulture);
            sed.Attributes["sepxOffset"] = sepxOffset.ToString(CultureInfo.InvariantCulture);
            if (sepxOffset >= 0)
            {
                if (sepxOffset > word.Length - 2)
                    throw new InvalidDataException("A section property block points outside WordDocument.");
                word.Position = sepxOffset;
                var propertyLength = unchecked((short)ReadU16(word));
                if (propertyLength < 0 || propertyLength > word.Length - sepxOffset - 2)
                    throw new InvalidDataException("A section property block has an invalid size.");
                var sepx = new DocStructureNode("Sepx", "SectionProperties", "WordDocument",
                    sepxOffset, propertyLength + 2L);
                sepx.Attributes["propertyBytes"] = propertyLength.ToString(CultureInfo.InvariantCulture);
                sed.Children.Add(sepx);
            }
            else if (sepxOffset != -1)
                throw new InvalidDataException("A section property offset is invalid.");
            plc.Children.Add(sed);
            previousCp = nextCp;
        }
        var main = structure.Parts.FirstOrDefault(x => x.Name == "Main");
        if (main != null && previousCp < main.CpEnd)
            throw new InvalidDataException("The final section CP precedes the end of the main document.");
        location.Children.Add(plc);
    }

    private static ushort ReadU16(Stream stream)
    {
        var bytes = ReadFully(stream, 2);
        return BinaryPrimitives.ReadUInt16LittleEndian(bytes);
    }

    private static uint ReadU32(Stream stream)
    {
        var bytes = ReadFully(stream, 4);
        return BinaryPrimitives.ReadUInt32LittleEndian(bytes);
    }

    private static byte[] ReadFully(Stream stream, int length)
    {
        var bytes = new byte[length];
        var read = 0;
        while (read < length)
        {
            var count = stream.Read(bytes, read, length - read);
            if (count == 0) throw new InvalidDataException("The section table ended unexpectedly.");
            read += count;
        }
        return bytes;
    }
}
