using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes shape anchors and their fixed SPA records.</summary>
internal static class DocShapeNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value,
            checked((int)location.Length.Value));
        if (bytes.Length < 4 || (bytes.Length - 4) % 30 != 0)
            throw new InvalidDataException("The shape-anchor PLC has an invalid length.");
        var count = (bytes.Length - 4) / 30;
        var cpBytes = checked((count + 1) * 4);
        var plc = new DocStructureNode("PlcfSpa", location.Name, location.StreamName,
            location.Offset, location.Length);
        plc.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        var storyStart = location.Name == "HeaderShapeAnchors"
            ? structure.Parts.FirstOrDefault(x => x.Name == "Headers")?.CpStart ?? 0u : 0u;
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4));
            var next = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i + 1) * 4));
            if (next <= cp) throw new InvalidDataException("Shape-anchor positions are not increasing.");
            var offset = checked(location.Offset.Value + cpBytes + i * 26L);
            var spa = new DocStructureNode("Spa", $"Shape{i}", location.StreamName, offset, 26);
            spa.Attributes["cp"] = cp.ToString(CultureInfo.InvariantCulture);
            spa.Attributes["globalCp"] = ((ulong)storyStart + cp).ToString(CultureInfo.InvariantCulture);
            spa.Attributes["shapeId"] = BinaryPrimitives.ReadUInt32LittleEndian(
                bytes.AsSpan(cpBytes + i * 26, 4)).ToString(CultureInfo.InvariantCulture);
            spa.Children.Add(new DocStructureNode("Rca", "Bounds", location.StreamName,
                offset + 4, 16));
            plc.Children.Add(spa);
        }
        location.Children.Add(plc);
    }
}
