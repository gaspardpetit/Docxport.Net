using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes textbox text and boundary PLCs with their source record ranges.</summary>
internal static class DocTextboxNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var boundary = location.Name.EndsWith("Boundaries", StringComparison.Ordinal);
        var header = location.Name.StartsWith("Header", StringComparison.Ordinal);
        var partName = header ? "HeaderTextboxes" : "Textboxes";
        var part = structure.Parts.FirstOrDefault(x => x.Name == partName);
        if (part == null) throw new InvalidDataException($"The {partName} PLC has no document part.");
        var recordSize = boundary ? 6 : 22;
        var length = checked((int)location.Length.Value);
        if (length < 4 || (length - 4) % (4 + recordSize) != 0)
            throw new InvalidDataException("The textbox PLC has an invalid size.");
        var count = (length - 4) / (4 + recordSize);
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        var kind = location.Name switch
        {
            "TextboxText" => "PlcftxbxTxt",
            "HeaderTextboxText" => "PlcfHdrtxbxTxt",
            "TextboxBoundaries" => "PlcfTxbxBkd",
            _ => "PlcfTxbxHdrBkd"
        };
        var plc = new DocStructureNode(kind, location.Name, location.StreamName,
            location.Offset, length);
        plc.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
        {
            var cpStart = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            var cpEnd = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i + 1) * 4, 4));
            if (cpStart >= cpEnd || cpEnd > part.Length)
                throw new InvalidDataException("A textbox CP range is invalid.");
            var recordOffset = (count + 1) * 4 + i * recordSize;
            var entry = new DocStructureNode(boundary ? "Tbkd" : "FTXBXS", $"Textbox{i}",
                location.StreamName, location.Offset.Value + recordOffset, recordSize);
            if (!boundary)
            {
                var reusable = i == count - 1 ||
                    (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(recordOffset + 8, 2)) & 1) != 0;
                entry.Children.Add(new DocStructureNode(reusable ? "FTXBXSReusable" : "FTXBXNonReusable",
                    reusable ? "ReusableState" : "TextboxState", location.StreamName,
                    entry.Offset, 8));
            }
            entry.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["cpStart"] = cpStart.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["cpEnd"] = cpEnd.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["globalCpStart"] = checked(part.CpStart + cpStart).ToString(CultureInfo.InvariantCulture);
            entry.Attributes["globalCpEnd"] = checked(part.CpStart + cpEnd).ToString(CultureInfo.InvariantCulture);
            if (boundary)
                entry.Attributes["textboxIndex"] = BinaryPrimitives.ReadInt16LittleEndian(
                    bytes.AsSpan(recordOffset, 2)).ToString(CultureInfo.InvariantCulture);
            plc.Children.Add(entry);
        }
        location.Children.Add(plc);
    }
}
