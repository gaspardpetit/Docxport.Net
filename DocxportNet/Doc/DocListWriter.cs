using System.Buffers.Binary;
using System.Text;

namespace DocxportNet.Doc;

internal static class DocListWriter
{
    public static (byte[] FixedDefinitions, byte[] Levels, byte[] Overrides)
        Write(DocListIndex lists, IReadOnlyDictionary<string, int> fontIndexes)
    {
        if (lists.Definitions.Count > short.MaxValue || lists.Instances.Count > short.MaxValue)
            throw new NotSupportedException("The DOC has too many lists.");
        var definitions = new byte[checked(2 + lists.Definitions.Count * 28)];
        BinaryPrimitives.WriteInt16LittleEndian(definitions, checked((short)lists.Definitions.Count));
        using var levels = new MemoryStream();
        var ids = new HashSet<int>();
        for (var i = 0; i < lists.Definitions.Count; i++)
        {
            var definition = lists.Definitions[i];
            if (!ids.Add(definition.ListId))
                throw new InvalidDataException("A DOC list ID is duplicated.");
            var position = 2 + i * 28;
            BinaryPrimitives.WriteInt32LittleEndian(definitions.AsSpan(position),
                definition.ListId);
            for (var style = 0; style < 9; style++)
                BinaryPrimitives.WriteUInt16LittleEndian(
                    definitions.AsSpan(position + 8 + style * 2), 0x0FFF);
            if (definition.Levels.Count == 1)
                definitions[position + 26] = 1;
            else if (definition.Levels.Count is < 1 or > 9)
                throw new NotSupportedException("A DOC list requires one or nine levels.");
            var count = definition.Levels.Count == 1 ? 1 : 9;
            for (var level = 0; level < count; level++)
            {
                var item = level < definition.Levels.Count
                    ? definition.Levels[level]
                    : new DocListLevel(level, 1, 0, $"%{level + 1}.", 0);
                if (item.Index != level)
                    throw new InvalidDataException("DOC list levels must be ordered.");
                var bytes = WriteLevel(item, fontIndexes);
                levels.Write(bytes, 0, bytes.Length);
            }
        }
        using var overrides = new MemoryStream();
        WriteI32(overrides, lists.Instances.Count);
        for (var i = 0; i < lists.Instances.Count; i++)
        {
            var instance = lists.Instances[i];
            if (instance.OverrideIndex != i + 1 || !ids.Contains(instance.ListId))
                throw new InvalidDataException("A DOC list instance has an invalid index or ID.");
            var count = instance.StartOverrides.Keys.Union(instance.FormattingOverrides.Keys)
                .Distinct().Count();
            if (count > 9)
                throw new InvalidDataException("A DOC list has too many level overrides.");
            WriteI32(overrides, instance.ListId);
            overrides.Write(new byte[8], 0, 8);
            overrides.WriteByte(checked((byte)count));
            overrides.WriteByte(0);
            overrides.WriteByte(0);
            overrides.WriteByte(0);
        }
        foreach (var instance in lists.Instances)
        {
            WriteI32(overrides, 0); // LFOData.cp.
            foreach (var levelIndex in instance.StartOverrides.Keys
                .Union(instance.FormattingOverrides.Keys).Distinct().OrderBy(x => x))
            {
                if (levelIndex is < 0 or > 8)
                    throw new InvalidDataException("A DOC list override has an invalid level.");
                var hasStart = instance.StartOverrides.TryGetValue(levelIndex, out var start);
                var hasFormatting = instance.FormattingOverrides.TryGetValue(levelIndex,
                    out var formatting);
                WriteI32(overrides, hasStart && !hasFormatting ? start : 0);
                overrides.WriteByte(checked((byte)(levelIndex |
                    (hasFormatting ? 0x20 : 0x10))));
                overrides.WriteByte(0);
                overrides.WriteByte(0);
                overrides.WriteByte(0);
                if (hasFormatting)
                {
                    var item = hasStart ? formatting! with { StartAt = start } : formatting!;
                    var bytes = WriteLevel(item, fontIndexes);
                    overrides.Write(bytes, 0, bytes.Length);
                }
            }
        }
        return (definitions, levels.ToArray(), overrides.ToArray());
    }

    private static byte[] WriteLevel(DocListLevel level,
        IReadOnlyDictionary<string, int> fontIndexes)
    {
        if (level.StartAt is < 0 or > 0x7FFF || level.FollowCharacter > 2 ||
            level.LabelJustification > 2 || level.RestartLimit is > 7)
            throw new InvalidDataException("A DOC list level has invalid values.");
        var header = new byte[28];
        BinaryPrimitives.WriteInt32LittleEndian(header, level.StartAt);
        header[4] = level.NumberFormat;
        header[5] = checked((byte)(level.LabelJustification |
            (level.RestartLimit != null ? 0x08 : 0)));
        header[15] = level.FollowCharacter;
        if (level.RestartLimit is byte restartLimit)
            header[26] = restartLimit;
        var characterProperties = level.LabelFormatting is { IsEmpty: false } formatting
            ? DocPlainTextWriter.EncodeCharacterProperties(formatting, fontIndexes)
            : Array.Empty<byte>();
        var paragraphProperties = level.ParagraphFormatting is { IsEmpty: false } paragraphFormatting
            ? paragraphFormatting.Encode()
            : Array.Empty<byte>();
        if (characterProperties.Length > byte.MaxValue)
            throw new NotSupportedException("A DOC list label has too many character properties.");
        if (paragraphProperties.Length > byte.MaxValue)
            throw new NotSupportedException("A DOC list level has too many paragraph properties.");
        header[24] = checked((byte)characterProperties.Length);
        header[25] = checked((byte)paragraphProperties.Length);
        var text = new StringBuilder();
        var placeholderCount = 0;
        for (var i = 0; i < level.NumberText.Length; i++)
        {
            if (level.NumberFormat is not (0x17 or 0xFF) &&
                level.NumberText[i] == '%' &&
                i + 1 < level.NumberText.Length &&
                level.NumberText[i + 1] is >= '1' and <= '9')
            {
                var referencedLevel = level.NumberText[++i] - '1';
                if (placeholderCount >= 9)
                    throw new NotSupportedException("A DOC list label has too many level placeholders.");
                header[6 + placeholderCount++] = checked((byte)(text.Length + 1));
                text.Append((char)referencedLevel);
            }
            else text.Append(level.NumberText[i]);
        }
        if (text.Length > ushort.MaxValue)
            throw new NotSupportedException("A DOC list label is too long.");
        using var output = new MemoryStream();
        output.Write(header, 0, header.Length);
        output.Write(paragraphProperties, 0, paragraphProperties.Length);
        output.Write(characterProperties, 0, characterProperties.Length);
        WriteU16(output, checked((ushort)text.Length));
        var encoded = Encoding.Unicode.GetBytes(text.ToString());
        output.Write(encoded, 0, encoded.Length);
        return output.ToArray();
    }

    private static void WriteI32(Stream stream, int value)
    {
        var bytes = new byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(bytes, value);
        stream.Write(bytes, 0, 4);
    }

    private static void WriteU16(Stream stream, ushort value)
    {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
    }
}
