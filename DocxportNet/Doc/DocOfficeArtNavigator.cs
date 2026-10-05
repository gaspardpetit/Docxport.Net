using System.Buffers.Binary;

namespace DocxportNet.Doc;

internal static class DocOfficeArtNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode content)
    {
        if (content.Children.Count != 0 || content.StreamName == null ||
            content.Offset == null || content.Length == null) return;
        var bytes = structure.ReadRange(content.StreamName, content.Offset.Value,
            checked((int)content.Length.Value));
        var cursor = 0;
        if (!TryRecord(bytes, 0, bytes.Length, out var groupEnd, out var groupType, out _) ||
            groupType != 0xF000) return;
        var group = new DocStructureNode("OfficeArtDggContainer", "DrawingGroup",
            content.StreamName, content.Offset, groupEnd);
        ExpandRecords(bytes, 8, groupEnd, group);
        content.Children.Add(group);
        cursor = groupEnd;
        while (cursor < bytes.Length)
        {
            if (bytes[cursor] is not (0 or 1) ||
                !TryRecord(bytes, cursor + 1, bytes.Length, out var end, out var type, out _) ||
                type != 0xF002) return;
            var drawing = new DocStructureNode("OfficeArtWordDrawing",
                bytes[cursor] == 0 ? "MainDrawing" : "HeaderDrawing",
                content.StreamName, content.Offset + cursor, end - cursor);
            var container = new DocStructureNode("OfficeArtDgContainer", "Drawing",
                content.StreamName, content.Offset + cursor + 1, end - cursor - 1);
            ExpandRecords(bytes, cursor + 9, end, container);
            drawing.Children.Add(container);
            content.Children.Add(drawing);
            cursor = end;
        }
    }

    private static void ExpandRecords(byte[] bytes, int begin, int end, DocStructureNode parent)
    {
        var cursor = begin;
        while (cursor < end)
        {
            if (!TryRecord(bytes, cursor, end, out var next, out var type, out var version)) return;
            var kind = type switch
            {
                0xF003 => "OfficeArtSpgrContainer",
                0xF004 => "OfficeArtSpContainer",
                0xF010 => "OfficeArtClientAnchor",
                0xF011 => "OfficeArtClientData",
                0xF00D => "OfficeArtClientTextbox",
                _ => "OfficeArtRecord"
            };
            var record = new DocStructureNode(kind, $"Record{parent.Children.Count}",
                parent.StreamName, parent.Offset + cursor - begin + 8, next - cursor);
            record.Attributes["recordType"] = $"0x{type:X4}";
            if (version == 15 && kind is ("OfficeArtSpgrContainer" or "OfficeArtSpContainer"))
                ExpandRecords(bytes, cursor + 8, next, record);
            parent.Children.Add(record);
            cursor = next;
        }
    }

    private static bool TryRecord(byte[] bytes, int cursor, int end,
        out int next, out ushort type, out int version)
    {
        next = cursor;
        type = 0;
        version = 0;
        if (cursor < 0 || end - cursor < 8) return false;
        var flags = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
        version = flags & 15;
        type = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 2));
        var size = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(cursor + 4));
        if (size > end - cursor - 8) return false;
        next = cursor + 8 + (int)size;
        return true;
    }
}
