using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes fixed list definitions and override headers; level payloads remain deferred.</summary>
internal static class DocListNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var definitions = location.Name == "ListDefinitions";
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value,
            checked((int)location.Length.Value));
        var headerSize = definitions ? 2 : 4;
        var recordSize = definitions ? 28 : 16;
        if (bytes.Length < headerSize) throw new InvalidDataException("The list table is truncated.");
        var count = definitions ? BinaryPrimitives.ReadInt16LittleEndian(bytes) :
            BinaryPrimitives.ReadInt32LittleEndian(bytes);
        if (count < 0 || count > (bytes.Length - headerSize) / recordSize)
            throw new InvalidDataException("The list table has an invalid count.");
        var table = new DocStructureNode(definitions ? "PlfLst" : "PlfLfo", location.Name,
            location.StreamName, location.Offset, location.Length);
        table.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        var levelCount = 0;
        for (var i = 0; i < count; i++)
        {
            var relative = headerSize + i * recordSize;
            var offset = location.Offset.Value + relative;
            var entry = new DocStructureNode(definitions ? "LSTF" : "LFO", $"List{i}",
                location.StreamName, offset, recordSize);
            entry.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["lsid"] = BinaryPrimitives.ReadInt32LittleEndian(
                bytes.AsSpan(relative, 4)).ToString(CultureInfo.InvariantCulture);
            if (definitions)
            {
                entry.Attributes["simple"] = (bytes[relative + 26] & 1) != 0 ? "true" : "false";
                levelCount += (bytes[relative + 26] & 1) != 0 ? 1 : 9;
                entry.Children.Add(DocTemplateCodeNavigator.Create(location.StreamName,
                    offset + 4, bytes.AsSpan(relative + 4, 4), "ListTemplateCode"));
                entry.Children.Add(new DocStructureNode("grfhic", "HtmlCompatibility", location.StreamName,
                    offset + 27, 1));
            }
            else
            {
                entry.Attributes["overrideLevelCount"] = bytes[relative + 12].ToString(CultureInfo.InvariantCulture);
                entry.Children.Add(new DocStructureNode("grfhic", "HtmlCompatibility", location.StreamName,
                    offset + 14, 1));
            }
            table.Children.Add(entry);
        }
        if (definitions)
        {
            var cursor = checked(location.Offset.Value + location.Length.Value);
            using var stream = structure.OpenStream(location.StreamName);
            var nextLocation = structure.Locations.Where(x => x.IsPresent &&
                x.StreamName == location.StreamName && x.Offset > cursor)
                .Select(x => (long)x.Offset).DefaultIfEmpty(stream.Length).Min();
            var limit = Math.Min(stream.Length, nextLocation);
            for (var i = 0; i < levelCount; i++)
            {
                var level = ReadLevel(structure, location.StreamName, cursor, limit, i);
                table.Children.Add(level);
                cursor += level.Length!.Value;
            }
            table.Attributes["levelCount"] = levelCount.ToString(CultureInfo.InvariantCulture);
        }
        else
        {
            var cursor = headerSize + count * recordSize;
            for (var i = 0; i < count; i++)
            {
                if (bytes.Length - cursor < 4)
                    throw new InvalidDataException("A list override data record is truncated.");
                var start = cursor;
                var data = new DocStructureNode("LFOData", $"OverrideData{i}", location.StreamName,
                    location.Offset.Value + start);
                data.Attributes["cp"] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(cursor, 4))
                    .ToString(CultureInfo.InvariantCulture);
                cursor += 4;
                var overrideCount = bytes[headerSize + i * recordSize + 12];
                for (var j = 0; j < overrideCount; j++)
                {
                    if (bytes.Length - cursor < 8)
                        throw new InvalidDataException("A list level override is truncated.");
                    var overrideStart = cursor;
                    var flags = bytes[cursor + 4];
                    cursor += 8;
                    var levelOverride = new DocStructureNode("LFOLVL", $"LevelOverride{j}",
                        location.StreamName, location.Offset.Value + overrideStart);
                    levelOverride.Attributes["level"] = (flags & 0x0F).ToString(CultureInfo.InvariantCulture);
                    levelOverride.Attributes["hasFormatting"] = (flags & 0x20) != 0 ? "true" : "false";
                    if ((flags & 0x20) != 0)
                    {
                        var level = ReadLevel(bytes, cursor, bytes.Length, location.StreamName,
                            location.Offset.Value, j);
                        levelOverride.Children.Add(level);
                        cursor += checked((int)level.Length!.Value);
                    }
                    levelOverride = Sized(levelOverride, cursor - overrideStart);
                    data.Children.Add(levelOverride);
                }
                data = Sized(data, cursor - start);
                table.Children.Add(data);
            }
            if (cursor != bytes.Length)
                throw new InvalidDataException("The list override table has trailing bytes.");
        }
        location.Children.Add(table);
    }

    private static DocStructureNode ReadLevel(DocStructure structure, string streamName,
        long offset, long limit, int index)
    {
        if (offset > limit - 28)
            throw new InvalidDataException("An appended list level is truncated.");
        var header = structure.ReadRange(streamName, offset, 28);
        var body = checked(28 + header[25] + header[24]);
        if (offset > limit - body - 2)
            throw new InvalidDataException("An appended list level has invalid property lengths.");
        var size = structure.ReadRange(streamName, offset + body, 2);
        var characters = BinaryPrimitives.ReadUInt16LittleEndian(size);
        var total = checked(body + 2 + characters * 2);
        if (offset > limit - total)
            throw new InvalidDataException("An appended list level has an invalid text length.");
        var bytes = structure.ReadRange(streamName, offset, total);
        return ReadLevel(bytes, 0, bytes.Length, streamName, offset, index);
    }

    private static DocStructureNode ReadLevel(byte[] bytes, int start, int limit,
        string streamName, long streamOffset, int index)
    {
        if (limit - start < 30)
            throw new InvalidDataException("A list level is truncated.");
        var paragraphBytes = bytes[start + 25];
        var characterBytes = bytes[start + 24];
        var stringOffset = checked(start + 28 + paragraphBytes + characterBytes);
        if (limit - stringOffset < 2)
            throw new InvalidDataException("A list level has invalid property lengths.");
        var characters = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(stringOffset, 2));
        var end = checked(stringOffset + 2 + characters * 2);
        if (end > limit) throw new InvalidDataException("A list level has an invalid text length.");
        var level = new DocStructureNode("LVL", $"Level{index}", streamName,
            streamOffset + start, end - start);
        var header = new DocStructureNode("LVLF", "LevelFields", streamName,
            streamOffset + start, 28);
        header.Attributes["paragraphPropertyBytes"] = paragraphBytes.ToString(CultureInfo.InvariantCulture);
        header.Attributes["characterPropertyBytes"] = characterBytes.ToString(CultureInfo.InvariantCulture);
        header.Children.Add(new DocStructureNode("grfhic", "HtmlCompatibility", streamName,
            streamOffset + start + 27, 1));
        level.Children.Add(header);
        level.Children.Add(new DocStructureNode("Xst", "NumberText", streamName,
            streamOffset + stringOffset, end - stringOffset));
        return level;
    }

    private static DocStructureNode Sized(DocStructureNode node, int length)
    {
        var sized = new DocStructureNode(node.Kind, node.Name, node.StreamName, node.Offset, length);
        foreach (var pair in node.Attributes) sized.Attributes[pair.Key] = pair.Value;
        foreach (var child in node.Children) sized.Children.Add(child);
        return sized;
    }
}
