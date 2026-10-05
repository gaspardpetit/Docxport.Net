using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Locates variable-length frame-set and list-style records.</summary>
internal static class DocFrameSetNavigator
{
    private static readonly string[] RecordKinds =
        ["", "DofrFsn", "DofrFsnp", "DofrFsnName", "DofrFsnFnm",
            "DofrFsnSpbd", "DofrRglstsf"];

    public static void ExpandListStyles(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length is not >= 4 || (node.Length - 4) % 4 != 0) return;
        var count = BinaryPrimitives.ReadInt32LittleEndian(
            structure.ReadRange(node.StreamName, node.Offset.Value, 4));
        if (count < 0 || 4L + count * 4 != node.Length) return;
        node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
            node.Children.Add(new DocStructureNode("Lstsf", $"ListStyle{i}",
                node.StreamName, node.Offset + 4 + i * 4L, 4));
    }

    public static void ExpandFrame(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null || node.Offset == null || node.Length < 36) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value, 36);
        var divider = new DocStructureNode("Fssd", "Divider", node.StreamName, node.Offset, 8);
        var units = new DocStructureNode("FssUnits", "Units", node.StreamName, node.Offset, 4);
        units.Attributes["value"] = BinaryPrimitives.ReadUInt32LittleEndian(bytes).ToString(CultureInfo.InvariantCulture);
        divider.Children.Add(units);
        node.Children.Add(divider);
        var kind = new DocStructureNode("Fsnk", "FrameKind", node.StreamName, node.Offset + 12, 4);
        kind.Attributes["value"] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(12)).ToString(CultureInfo.InvariantCulture);
        node.Children.Add(kind);
        var scroll = new DocStructureNode("IScrollType", "ScrollBars", node.StreamName, node.Offset + 24, 4);
        scroll.Attributes["value"] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(24)).ToString(CultureInfo.InvariantCulture);
        node.Children.Add(scroll);
    }

    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length is not > 0) return;
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value,
            checked((int)location.Length.Value));
        var cursor = 0;
        while (cursor < bytes.Length)
        {
            if (bytes.Length - cursor < 8) return;
            var length = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(cursor));
            var type = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(cursor + 4));
            if (length < 8 || length > bytes.Length - cursor || type >= RecordKinds.Length) return;
            var header = new DocStructureNode("Dofrh", $"Record{location.Children.Count}",
                location.StreamName, location.Offset + cursor, length);
            var tag = new DocStructureNode("Dofrt", "RecordType", location.StreamName,
                location.Offset + cursor + 4, 4);
            tag.Attributes["value"] = type.ToString(CultureInfo.InvariantCulture);
            header.Children.Add(tag);
            if (type != 0)
            {
                var payload = new DocStructureNode("Dofr", "Payload", location.StreamName,
                    location.Offset + cursor + 8, length - 8);
                payload.Children.Add(new DocStructureNode(RecordKinds[type], "Value",
                    location.StreamName, payload.Offset, payload.Length));
                header.Children.Add(payload);
            }
            location.Children.Add(header);
            cursor += checked((int)length);
        }
    }
}
