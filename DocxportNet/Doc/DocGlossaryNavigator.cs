using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes CP-only glossary item ranges.</summary>
internal static class DocGlossaryNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var length = checked((int)location.Length.Value);
        if (length < 8 || length % 4 != 0)
            throw new InvalidDataException("The glossary PLC has an invalid length.");
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        var count = length / 4 - 2;
        var plc = new DocStructureNode("PlcfGlsy", "GlossaryRanges", location.StreamName,
            location.Offset, length);
        plc.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 1; i < length / 4; i++)
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4)) <=
                BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i - 1) * 4, 4)))
                throw new InvalidDataException("Glossary CPs are not increasing.");
        for (var i = 0; i < count; i++)
        {
            var start = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            var end = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i + 1) * 4, 4));
            var item = new DocStructureNode("GlossaryItem", $"Item{i}", location.StreamName,
                location.Offset.Value + i * 4L, 8);
            item.Attributes["cpStart"] = start.ToString(CultureInfo.InvariantCulture);
            item.Attributes["cpEnd"] = end.ToString(CultureInfo.InvariantCulture);
            plc.Children.Add(item);
        }
        location.Children.Add(plc);
    }
}
