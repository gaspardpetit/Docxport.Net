using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes the optional, deprecated table-character cache.</summary>
internal static class DocTableCharacterNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value,
            checked((int)location.Length.Value));
        if (bytes.Length < 4 || (bytes.Length - 4) % 8 != 0)
            throw new InvalidDataException("The table-character PLC has an invalid length.");
        var count = (bytes.Length - 4) / 8;
        var records = (count + 1) * 4;
        var plc = new DocStructureNode("PlcfTch", "TableCharacterCache", location.StreamName,
            location.Offset, location.Length);
        plc.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            var next = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i + 1) * 4, 4));
            if (next <= cp) throw new InvalidDataException("Table-character CPs are not increasing.");
            var record = new DocStructureNode("Tch", $"Range{i}", location.StreamName,
                location.Offset.Value + records + i * 4L, 4);
            record.Attributes["cpStart"] = cp.ToString(CultureInfo.InvariantCulture);
            record.Attributes["cpEnd"] = next.ToString(CultureInfo.InvariantCulture);
            plc.Children.Add(record);
        }
        location.Children.Add(plc);
    }
}
