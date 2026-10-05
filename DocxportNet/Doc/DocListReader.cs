using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace DocxportNet.Doc;

public sealed record DocListLevel(int Index, int StartAt, byte NumberFormat,
    string NumberText, byte FollowCharacter,
    DocCharacterFormatting? LabelFormatting = null,
    DocParagraphFormatting? ParagraphFormatting = null,
    byte LabelJustification = 0,
    byte? RestartLimit = null);

public sealed record DocListDefinition(int ListId, IReadOnlyList<DocListLevel> Levels);

public sealed record DocListInstance(short OverrideIndex, int ListId,
    IReadOnlyDictionary<int, int> StartOverrides,
    IReadOnlyDictionary<int, DocListLevel> FormattingOverrides);

public sealed record DocListIndex(IReadOnlyList<DocListDefinition> Definitions,
    IReadOnlyList<DocListInstance> Instances);

internal static class DocListReader
{
    public static DocListIndex Read(DocTextIndex index)
    {
        var definitions = Expand(index.Structure, "ListDefinitions");
        var overrides = Expand(index.Structure, "ListOverrides");
        if (definitions == null && overrides == null)
            return new DocListIndex([], []);
        if (definitions == null || overrides == null)
            throw new InvalidDataException("The DOC list tables are incomplete.");
        var listRecords = definitions.Children.Where(x => x.Kind == "LSTF").ToArray();
        var levelRecords = definitions.Children.Where(x => x.Kind == "LVL").ToArray();
        var levelCursor = 0;
        var parsedDefinitions = new List<DocListDefinition>();
        foreach (var record in listRecords)
        {
            var count = record.Attributes["simple"] == "true" ? 1 : 9;
            if (levelCursor + count > levelRecords.Length)
                throw new InvalidDataException("The DOC list has missing levels.");
            var levels = new DocListLevel[count];
            for (var i = 0; i < count; i++)
                levels[i] = ReadLevel(index, levelRecords[levelCursor++], i);
            parsedDefinitions.Add(new DocListDefinition(
                int.Parse(record.Attributes["lsid"], CultureInfo.InvariantCulture), levels));
        }
        var instanceRecords = overrides.Children.Where(x => x.Kind == "LFO").ToArray();
        var dataRecords = overrides.Children.Where(x => x.Kind == "LFOData").ToArray();
        if (instanceRecords.Length != dataRecords.Length)
            throw new InvalidDataException("The DOC list override table has inconsistent counts.");
        var instances = new DocListInstance[instanceRecords.Length];
        for (var i = 0; i < instances.Length; i++)
        {
            var starts = new Dictionary<int, int>();
            var formats = new Dictionary<int, DocListLevel>();
            foreach (var levelOverride in dataRecords[i].Children.Where(x => x.Kind == "LFOLVL"))
            {
                var bytes = index.Structure.ReadRange(levelOverride.StreamName!,
                    levelOverride.Offset!.Value, 8);
                var level = bytes[4] & 0x0F;
                if ((bytes[4] & 0x10) != 0 && (bytes[4] & 0x20) == 0)
                    starts[level] = BinaryPrimitives.ReadInt32LittleEndian(bytes);
                if ((bytes[4] & 0x20) != 0)
                {
                    var nested = levelOverride.Children.Single(x => x.Kind == "LVL");
                    formats[level] = ReadLevel(index, nested, level);
                }
            }
            instances[i] = new DocListInstance(checked((short)(i + 1)),
                int.Parse(instanceRecords[i].Attributes["lsid"], CultureInfo.InvariantCulture),
                starts, formats);
        }
        return new DocListIndex(parsedDefinitions, instances);
    }

    private static DocListLevel ReadLevel(DocTextIndex index, DocStructureNode node,
        int level)
    {
        var bytes = index.Structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length!.Value));
        if (bytes.Length < 30)
            throw new InvalidDataException("A DOC list level is truncated.");
        DocParagraphFormatting? paragraphFormatting = null;
        if (bytes[25] != 0)
        {
            if (28 + bytes[25] > bytes.Length)
                throw new InvalidDataException("A DOC list paragraph property range is truncated.");
            paragraphFormatting = DocParagraphFormatting.Parse(bytes.AsSpan(28, bytes[25]));
        }
        DocCharacterFormatting? labelFormatting = null;
        if (bytes[24] != 0)
        {
            var characterStart = checked(28 + bytes[25]);
            if (characterStart + bytes[24] > bytes.Length)
                throw new InvalidDataException("A DOC list label property range is truncated.");
            labelFormatting = DocCharacterFormattingReader.ParseSprms(
                bytes.AsSpan(characterStart, bytes[24]));
            if (labelFormatting.AsciiFontIndex != null ||
                labelFormatting.EastAsiaFontIndex != null ||
                labelFormatting.HighAnsiFontIndex != null)
                labelFormatting = DocCharacterFormattingReader.ResolveFonts(
                    labelFormatting, index.Fonts);
        }
        var textOffset = checked(28 + bytes[24] + bytes[25]);
        if (textOffset + 2 > bytes.Length)
            throw new InvalidDataException("A DOC list level has no number text.");
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(textOffset));
        if (textOffset + 2 + count * 2 > bytes.Length)
            throw new InvalidDataException("A DOC list level has truncated number text.");
        var text = Encoding.Unicode.GetString(bytes, textOffset + 2, count * 2);
        if (bytes[4] != 0x17)
        {
            var placeholders = new bool[text.Length];
            for (var i = 0; i < 9; i++)
            {
                var position = bytes[6 + i] - 1;
                if (position >= 0 && position < text.Length && text[position] <= 8)
                    placeholders[position] = true;
            }
            var builder = new StringBuilder();
            for (var i = 0; i < text.Length; i++)
            {
                if (placeholders[i])
                {
                    builder.Append('%');
                    builder.Append((char)('1' + text[i]));
                }
                else builder.Append(text[i]);
            }
            text = builder.ToString();
        }
        return new DocListLevel(level,
            BinaryPrimitives.ReadInt32LittleEndian(bytes), bytes[4], text, bytes[15],
            labelFormatting, paragraphFormatting, (byte)(bytes[5] & 3),
            (bytes[5] & 0x08) != 0 ? bytes[26] : null);
    }

    private static DocStructureNode? Expand(DocStructure structure, string name)
    {
        var location = structure.FindLocation(name);
        if (location?.IsPresent != true) return null;
        var node = new DocStructureNode("FibLocation", name, location.StreamName,
            location.Offset, location.Length);
        DocListNavigator.Expand(structure, node);
        return node.Children.Single();
    }
}
