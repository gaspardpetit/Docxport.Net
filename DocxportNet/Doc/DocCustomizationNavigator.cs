using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes the fixed-size command and key-map records in a Tcg.</summary>
internal static class DocCustomizationNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length == null) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value,
            checked((int)node.Length.Value));
        if (node.Kind == "Tcg")
        {
            if (bytes.Length > 1 && bytes[0] == 255)
                node.Children.Add(new DocStructureNode("Tcg255", "Customizations",
                    node.StreamName, node.Offset + 1, bytes.Length - 1));
            return;
        }
        if (node.Kind == "Tcg255")
        {
            var cursor = 0;
            while (cursor < bytes.Length)
            {
                if (bytes[cursor] == 0x40) break;
                if (bytes[cursor] is 0x10 or 0x11)
                {
                    var segmentLength = bytes[cursor] == 0x10
                        ? StringTableLength(bytes.AsSpan(cursor))
                        : MacroNamesLength(bytes.AsSpan(cursor));
                    if (segmentLength <= 0) break;
                    node.Children.Add(new DocStructureNode(bytes[cursor] == 0x10
                            ? "TcgSttbf" : "MacroNames", $"Customization{node.Children.Count}",
                        node.StreamName, node.Offset + cursor, segmentLength));
                    cursor += segmentLength;
                    continue;
                }
                if (bytes[cursor] == 0x12)
                {
                    var segmentLength = ToolbarLength(bytes.AsSpan(cursor));
                    if (segmentLength <= 0) break;
                    node.Children.Add(new DocStructureNode("CTBWRAPPER", $"Toolbar{node.Children.Count}",
                        node.StreamName, node.Offset + cursor, segmentLength));
                    cursor += segmentLength;
                    continue;
                }
                var (kind, size) = bytes[cursor] switch
                {
                    1 => ("PlfMcd", 24),
                    2 => ("PlfAcd", 4),
                    3 or 4 => ("PlfKme", 14),
                    _ => ("", 0)
                };
                if (size == 0 || bytes.Length - cursor < 5) break;
                var count = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(cursor + 1));
                if (count < 0 || count > (bytes.Length - cursor - 5) / size) break;
                var length = checked(5 + count * size);
                node.Children.Add(new DocStructureNode(kind, $"Customization{node.Children.Count}",
                    node.StreamName, node.Offset + cursor, length));
                cursor += length;
            }
            return;
        }
        if (node.Kind == "TcgSttbf" && StringTableLength(bytes) == bytes.Length)
        {
            node.Children.Add(new DocStructureNode("TcgSttbfCore", "CommandStrings",
                node.StreamName, node.Offset + 1, bytes.Length - 1));
            return;
        }
        if (node.Kind == "CTBWRAPPER")
        {
            var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(10));
            var cursor = 16 + BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(12));
            for (var i = 0; i < count; i++)
            {
                var length = CustomizationLength(bytes.AsSpan(cursor));
                if (length <= 0) return;
                node.Children.Add(new DocStructureNode("Customization", $"Customization{i}",
                    node.StreamName, node.Offset + cursor, length));
                cursor += length;
            }
            node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
            return;
        }
        if (node.Kind == "Customization")
        {
            var toolbarId = BinaryPrimitives.ReadUInt32LittleEndian(bytes);
            if (toolbarId == 0)
                node.Children.Add(new DocStructureNode("CTB", "CustomToolbar", node.StreamName,
                    node.Offset + 8, bytes.Length - 8));
            else
            {
                var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(6));
                for (var i = 0; i < count; i++)
                    node.Children.Add(new DocStructureNode("TBDelta", $"Delta{i}", node.StreamName,
                        node.Offset + 8 + i * 18L, 18));
            }
            return;
        }
        if (node.Kind == "CTB")
        {
            var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
            var cursor = 2 + chars * 2;
            var dataLength = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(cursor));
            var controlCount = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(cursor + dataLength));
            node.Attributes["controlCount"] = controlCount.ToString(CultureInfo.InvariantCulture);
            return;
        }
        if (node.Kind == "TBDelta")
        {
            var flags = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(14));
            var length = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(16));
            var offset = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(10));
            var table = structure.Root.Children.FirstOrDefault(x => x.Kind == "Stream" &&
                x.Name == node.StreamName);
            if ((flags & 1) != 0 && length >= 11 && table?.Length is long tableLength &&
                offset <= tableLength - length)
                node.Children.Add(new DocStructureNode("TBC", "ToolbarControl", node.StreamName,
                    offset, length));
            return;
        }
        if (node.Kind == "TcgSttbfCore" && bytes.Length >= 6)
        {
            node.Attributes["entryCount"] = BinaryPrimitives.ReadUInt16LittleEndian(
                bytes.AsSpan(2)).ToString(CultureInfo.InvariantCulture);
            return;
        }
        if (node.Kind == "MacroNames" && MacroNamesLength(bytes) == bytes.Length)
        {
            var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(1));
            var cursor = 3;
            for (var i = 0; i < count; i++)
            {
                var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 2));
                var length = 6 + chars * 2;
                var entry = new DocStructureNode("MacroName", $"Macro{i}", node.StreamName,
                    node.Offset + cursor, length);
                entry.Children.Add(new DocStructureNode("Xstz", "Name", node.StreamName,
                    node.Offset + cursor + 2, length - 2));
                node.Children.Add(entry);
                cursor += length;
            }
            node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
            return;
        }
        if (node.Kind is "PlfMcd" or "PlfAcd" or "PlfKme" && bytes.Length >= 5)
        {
            var (kind, size) = node.Kind switch
            {
                "PlfMcd" => ("Mcd", 24),
                "PlfAcd" => ("Acd", 4),
                _ => ("Kme", 14)
            };
            var count = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(1));
            if (count < 0 || 5L + count * size != bytes.Length) return;
            node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
            for (var i = 0; i < count; i++)
                node.Children.Add(new DocStructureNode(kind, $"Entry{i}", node.StreamName,
                    node.Offset + 5 + i * (long)size, size));
            return;
        }
        if (node.Kind == "Kme" && bytes.Length == 14)
        {
            node.Children.Add(new DocStructureNode("Kcm", "PrimaryKey", node.StreamName,
                node.Offset + 4, 2));
            node.Children.Add(new DocStructureNode("Kcm", "SecondaryKey", node.StreamName,
                node.Offset + 6, 2));
            var action = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(8));
            var kt = new DocStructureNode("Kt", "ActionKind", node.StreamName,
                node.Offset + 8, 2);
            kt.Attributes["value"] = action.ToString(CultureInfo.InvariantCulture);
            node.Children.Add(kt);
            if (action == 0)
                node.Children.Add(new DocStructureNode("Cid", "Command", node.StreamName,
                    node.Offset + 10, 4));
        }
        else if (node.Kind == "Acd" && bytes.Length == 4)
        {
            node.Children.Add(new DocStructureNode("Fci", "Command", node.StreamName,
                node.Offset + 2, 2));
        }
        else if (node.Kind == "Cid" && bytes.Length == 4)
        {
            var commandType = bytes[0] & 7;
            var type = new DocStructureNode("Cmt", "CommandType", node.StreamName,
                node.Offset, 1);
            type.Attributes["value"] = commandType.ToString(CultureInfo.InvariantCulture);
            node.Children.Add(type);
            var variant = commandType switch
            {
                1 => "CidFci",
                2 => "CidMacro",
                3 => "CidAllocated",
                _ => null
            };
            if (variant != null)
                node.Children.Add(new DocStructureNode(variant, "CommandValue",
                    node.StreamName, node.Offset, 4));
        }
        else if (node.Kind == "CidFci" && bytes.Length == 4)
        {
            var command = new DocStructureNode("Fci", "BuiltInCommand",
                node.StreamName, node.Offset, 2);
            command.Attributes["value"] = (BinaryPrimitives.ReadUInt16LittleEndian(bytes) >> 3)
                .ToString(CultureInfo.InvariantCulture);
            node.Children.Add(command);
        }
    }

    private static int ToolbarLength(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 24 || bytes[0] != 0x12 || bytes[3] != 7 ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(4)) != 6 ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6)) != 12 ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8)) != 18) return -1;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(10));
        var controls = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice(12));
        if (count == 0 || controls < 0 || controls > bytes.Length - 16) return -1;
        var cursor = 16 + controls;
        for (var i = 0; i < count; i++)
        {
            var size = CustomizationLength(bytes.Slice(cursor));
            if (size <= 0) return -1;
            cursor += size;
        }
        return cursor;
    }

    private static int CustomizationLength(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 8 || BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(4)) != 0)
            return -1;
        var toolbarId = BinaryPrimitives.ReadUInt32LittleEndian(bytes);
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(6));
        if (toolbarId != 0)
            return count <= (bytes.Length - 8) / 18 ? 8 + count * 18 : -1;
        if (count != 0 || bytes.Length < 14) return -1;
        var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(8));
        var nameEnd = 10L + chars * 2L;
        if (nameEnd + 4 > bytes.Length) return -1;
        var dataLength = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice((int)nameEnd));
        if (dataLength < 4 || nameEnd + dataLength + 4 > bytes.Length) return -1;
        var controls = BinaryPrimitives.ReadInt32LittleEndian(bytes.Slice((int)(nameEnd + dataLength)));
        if (controls != 0) return -1;
        return checked((int)(nameEnd + dataLength + 4));
    }

    private static int StringTableLength(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 7 || bytes[0] != 0x10 ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(1)) != 0xFFFF ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(5)) != 2) return -1;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(3));
        var cursor = 7;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 2) return -1;
            var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(cursor));
            var length = 4L + 2L * chars;
            if (length > bytes.Length - cursor) return -1;
            cursor += (int)length;
        }
        return cursor;
    }

    private static int MacroNamesLength(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 3 || bytes[0] != 0x11) return -1;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(1));
        var cursor = 3;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 4) return -1;
            var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(cursor + 2));
            var length = 6L + 2L * chars;
            if (chars > 255 || length > bytes.Length - cursor) return -1;
            cursor += (int)length;
        }
        return cursor;
    }
}
