using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

internal static class DocSdtNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value,
            checked((int)location.Length.Value));
        if (location.Name == "SdtBookmarkNames") ExpandBookmarks(location, bytes);
        else ExpandSchemas(location, bytes);
    }

    private static void ExpandBookmarks(DocStructureNode location, byte[] bytes)
    {
        if (bytes.Length < 8 || U16(bytes, 0) != 0xFFFF || U16(bytes, 6) != 0) return;
        var count = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(2));
        if (count < 0 || count > (bytes.Length - 8) / 26) return;
        var table = new DocStructureNode("SttbfBkmkSdt", "StructuredTagBookmarks",
            location.StreamName, location.Offset, location.Length);
        var cursor = 8;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 26 || U16(bytes, cursor) != 12) return;
            var begin = cursor + 2;
            var attributes = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(begin + 16));
            var placeholderBytes = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(begin + 20));
            if (attributes > (bytes.Length - begin - 24) / 12 ||
                placeholderBytes > bytes.Length - begin - 24 || placeholderBytes % 2 != 0) return;
            var item = new DocStructureNode("SDTI", $"Tag{i}", location.StreamName,
                location.Offset + begin, 24);
            item.Children.Add(new DocStructureNode("TIQ", "TagIdentity", location.StreamName,
                item.Offset + 4, 8));
            item.Children.Add(new DocStructureNode("SDTT", "TagType", location.StreamName,
                item.Offset + 12, 4));
            var next = begin + 24;
            for (var j = 0; j < attributes; j++)
            {
                if (bytes.Length - next < 10) return;
                var chars = U16(bytes, next + 8);
                var size = 10L + 2L * (chars + 1);
                if (size > bytes.Length - next) return;
                var attribute = new DocStructureNode("FSDAP", $"Attribute{j}", location.StreamName,
                    location.Offset + next, size);
                attribute.Children.Add(new DocStructureNode("TIQ", "AttributeIdentity",
                    location.StreamName, attribute.Offset, 8));
                item.Children.Add(attribute);
                next += (int)size;
            }
            if (placeholderBytes > bytes.Length - next) return;
            item = CopyWithLength(item, next + (int)placeholderBytes - begin);
            table.Children.Add(item);
            cursor = next + (int)placeholderBytes;
        }
        if (cursor != bytes.Length) return;
        table.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        location.Children.Add(table);
    }

    private static DocStructureNode CopyWithLength(DocStructureNode source, long length)
    {
        var copy = new DocStructureNode(source.Kind, source.Name, source.StreamName, source.Offset, length);
        foreach (var child in source.Children) copy.Children.Add(child);
        return copy;
    }

    private static void ExpandSchemas(DocStructureNode location, byte[] bytes)
    {
        if (bytes.Length < 4) return;
        var count = BinaryPrimitives.ReadInt32LittleEndian(bytes);
        if (count < 0 || count > (bytes.Length - 4) / 16) return;
        var schema = new DocStructureNode("Hplxsdr", "SchemaReferences", location.StreamName,
            location.Offset, location.Length);
        var cursor = 4;
        for (var i = 0; i < count; i++)
        {
            var begin = cursor;
            for (var field = 0; field < 2; field++)
            {
                if (bytes.Length - cursor < 2) return;
                var chars = U16(bytes, cursor);
                var size = 2L + chars * 2L;
                if (size > bytes.Length - cursor) return;
                cursor += (int)size;
            }
            for (var field = 0; field < 2; field++)
            {
                var size = WideCountStringTableLength(bytes.AsSpan(cursor));
                if (size < 0) return;
                cursor += size;
            }
            schema.Children.Add(new DocStructureNode("XSDR", $"Schema{i}",
                location.StreamName, location.Offset + begin, cursor - begin));
        }
        if (cursor != bytes.Length) return;
        schema.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        location.Children.Add(schema);
    }

    private static int WideCountStringTableLength(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 8 || U16(bytes, 0) != 0xFFFF) return -1;
        var count = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(2));
        var extra = U16(bytes, 6);
        if (count < 0 || count > (bytes.Length - 8) / 2) return -1;
        var cursor = 8;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 2) return -1;
            var size = 2L + U16(bytes, cursor) * 2L + extra;
            if (size > bytes.Length - cursor) return -1;
            cursor += (int)size;
        }
        return cursor;
    }

    private static ushort U16(ReadOnlySpan<byte> bytes, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(offset, 2));
}
