using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Locates header stories and note references/text without decoding their contents.</summary>
internal static class DocStoryPlcNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;

        var name = location.Name;
        var references = name.EndsWith("References", StringComparison.Ordinal);
        var partName = name switch
        {
            "HeadersAndFooters" => "Headers",
            "FootnoteText" => "Footnotes",
            "EndnoteText" => "Endnotes",
            "CommentText" => "Comments",
            _ => "Main"
        };
        var kind = name switch
        {
            "HeadersAndFooters" => "Plcfhdd",
            "FootnoteReferences" => "PlcffndRef",
            "FootnoteText" => "PlcffndTxt",
            "EndnoteReferences" => "PlcfendRef",
            "CommentReferences" => "PlcfandRef",
            "CommentText" => "PlcfandTxt",
            _ => "PlcfendTxt"
        };
        var part = structure.Parts.FirstOrDefault(x => x.Name == partName);
        if (part == null) throw new InvalidDataException($"{kind} has no {partName} document part.");

        var length = location.Length.Value;
        var recordSize = name == "CommentReferences" ? 30 : 2;
        var count = references
            ? length >= 4 && (length - 4) % (4 + recordSize) == 0
                ? checked((int)((length - 4) / (4 + recordSize))) : -1
            : length >= 8 && length % 4 == 0 ? checked((int)(length / 4 - 2)) : -1;
        if (count < 0) throw new InvalidDataException($"{kind} has an invalid size.");
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, checked((int)length));
        var plc = new DocStructureNode(kind, name, location.StreamName, location.Offset, length);
        plc.Attributes[references ? "referenceCount" : "storyCount"] = count.ToString(CultureInfo.InvariantCulture);
        var cpCount = references ? count + 1 : count + 2;
        var cp = new uint[cpCount];
        for (var i = 0; i < cpCount; i++)
            cp[i] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));

        if (references)
        {
            for (var i = 0; i < count; i++)
            {
                if (cp[i] >= part.Length || (i > 0 && cp[i] <= cp[i - 1]))
                    throw new InvalidDataException($"{kind} has an invalid reference CP.");
                var recordOffset = location.Offset.Value + (count + 1L) * 4 + i * recordSize;
                var node = new DocStructureNode(kind + "Entry", $"Reference{i}", location.StreamName,
                    recordOffset, recordSize);
                node.Attributes["cp"] = cp[i].ToString(CultureInfo.InvariantCulture);
                node.Attributes["globalCp"] = checked(part.CpStart + cp[i]).ToString(CultureInfo.InvariantCulture);
                node.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
                if (name == "CommentReferences")
                    node.Children.Add(new DocStructureNode("ATRDPre10", "CommentRecord", location.StreamName,
                        recordOffset, recordSize));
                else
                    node.Attributes["numberingIndex"] = BinaryPrimitives.ReadUInt16LittleEndian(
                        bytes.AsSpan(checked((int)(recordOffset - location.Offset.Value)), 2)).ToString(CultureInfo.InvariantCulture);
                plc.Children.Add(node);
            }
        }
        else
        {
            if (part.Length == 0 || cp[count] != part.Length - 1)
                throw new InvalidDataException($"{kind} has an invalid final story CP.");
            for (var i = 0; i < count; i++)
            {
                if (cp[i] > cp[i + 1] || cp[i + 1] > part.Length - 1 ||
                    (i > 0 && cp[i] < cp[i - 1]) ||
                    (name != "HeadersAndFooters" && cp[i] == cp[i + 1]))
                    throw new InvalidDataException($"{kind} has an invalid story CP range.");
                var node = new DocStructureNode(kind + "Entry", $"Story{i}", location.StreamName,
                    location.Offset.Value + i * 4L, 8);
                node.Attributes["cpStart"] = cp[i].ToString(CultureInfo.InvariantCulture);
                node.Attributes["cpEnd"] = cp[i + 1].ToString(CultureInfo.InvariantCulture);
                node.Attributes["globalCpStart"] = checked(part.CpStart + cp[i]).ToString(CultureInfo.InvariantCulture);
                node.Attributes["globalCpEnd"] = checked(part.CpStart + cp[i + 1]).ToString(CultureInfo.InvariantCulture);
                node.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
                plc.Children.Add(node);
            }
        }
        location.Children.Add(plc);
    }
}
