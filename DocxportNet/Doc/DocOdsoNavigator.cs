using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes bounded mail-merge data-source properties in fcODSO.</summary>
internal static class DocOdsoNavigator
{
    public static void ExpandList(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length is not >= 12 ||
            node.Kind is not ("RecipientInfo" or "FieldMapInfo")) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value,
            checked((int)node.Length.Value));
        if (BinaryPrimitives.ReadUInt16LittleEndian(bytes) != 0 ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2)) != 4 ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(8)) != 1) return;
        var count = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(4));
        uint listLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(10));
        var cursor = 12;
        if (listLength == 0xFFFF)
        {
            if (bytes.Length < 16) return;
            listLength = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(12));
            cursor = 16;
        }
        if (listLength > bytes.Length - cursor || count > listLength / 4) return;
        var end = cursor + (int)listLength;
        var groups = new List<DocStructureNode>();
        var fieldMap = node.Kind == "FieldMapInfo";
        for (var i = 0; i < count; i++)
        {
            var begin = cursor;
            var entries = new List<DocStructureNode>();
            while (end - cursor >= 4)
            {
                var id = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
                var length = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 2));
                if (id == 0 && length == 0)
                {
                    entries.Add(new DocStructureNode(fieldMap ? "FieldMapTerminator" :
                        "RecipientTerminator", "End", node.StreamName, node.Offset + cursor, 4));
                    cursor += 4;
                    break;
                }
                if (id is < 1 or > 4 || length > end - cursor - 4) return;
                var item = new DocStructureNode(fieldMap ? "FieldMapDataItem" :
                    "RecipientDataItem", $"Data{id}", node.StreamName,
                    node.Offset + cursor, 4 + length);
                item.Attributes["id"] = id.ToString(CultureInfo.InvariantCulture);
                entries.Add(item);
                cursor += 4 + length;
            }
            if (entries.Count == 0 || entries[entries.Count - 1].Kind is not
                ("FieldMapTerminator" or "RecipientTerminator")) return;
            var group = new DocStructureNode(fieldMap ? "FieldMapBase" : "RecipientBase",
                $"Entry{i}", node.StreamName, node.Offset + begin, cursor - begin);
            foreach (var entry in entries) group.Children.Add(entry);
            groups.Add(group);
        }
        if (cursor != end) return;
        foreach (var group in groups) node.Children.Add(group);
    }

    public static void Expand(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length == null) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value,
            checked((int)node.Length.Value));
        var cursor = 0;
        while (bytes.Length - cursor >= 4)
        {
            var id = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
            var cb = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 2));
            var large = cb == 0xFFFF;
            if (large && bytes.Length - cursor < 8) break;
            var payloadLength = large
                ? BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(cursor + 4))
                : cb;
            var prefix = large ? 8 : 4;
            if (payloadLength > bytes.Length - cursor - prefix) break;
            var property = new DocStructureNode("ODSOPropertyBase", $"Property{node.Children.Count}",
                node.StreamName, node.Offset + cursor, prefix + payloadLength);
            property.Attributes["id"] = $"0x{id:X4}";
            var value = new DocStructureNode(large ? "ODSOPropertyLarge" : "ODSOPropertyStandard",
                "Value", node.StreamName, node.Offset + cursor + 4,
                payloadLength + (large ? 4 : 0));
            property.Children.Add(value);
            var start = cursor + prefix;
            if (id == 0x13)
            {
                var item = start;
                while (start + (long)payloadLength - item >= 16)
                {
                    var size = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(item));
                    if (size < 16 || size > start + (long)payloadLength - item) break;
                    value.Children.Add(new DocStructureNode("FilterDataItem", "Filter",
                        node.StreamName, node.Offset + item, size));
                    item += (int)size;
                }
            }
            else if (id == 0x14 && payloadLength <= 24 && payloadLength % 8 == 0)
            {
                for (var i = 0; i < payloadLength / 8; i++)
                    value.Children.Add(new DocStructureNode("SortColumnAndDirection", $"Sort{i}",
                        node.StreamName, node.Offset + start + i * 8, 8));
            }
            else if (id is 0x15 or 0x16 && payloadLength >= 12 &&
                BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(start)) == 0 &&
                BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(start + 2)) == 4 &&
                BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(start + 8)) == 1)
            {
                var kind = id == 0x15 ? "RecipientInfo" : "FieldMapInfo";
                var info = new DocStructureNode(kind, "Mapping", node.StreamName,
                    node.Offset + start, payloadLength);
                info.Attributes["entryCount"] = BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(start + 4)).ToString(CultureInfo.InvariantCulture);
                value.Children.Add(info);
            }
            node.Children.Add(property);
            cursor += checked(prefix + (int)payloadLength);
        }
        node.Attributes["parsedBytes"] = cursor.ToString(CultureInfo.InvariantCulture);
    }
}
