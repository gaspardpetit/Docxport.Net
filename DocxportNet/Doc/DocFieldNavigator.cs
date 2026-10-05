using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes field characters in each document story without interpreting field instructions.</summary>
internal static class DocFieldNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var partName = location.Name switch
        {
            "MainFields" => "Main",
            "HeaderFields" => "Headers",
            "FootnoteFields" => "Footnotes",
            "CommentFields" => "Comments",
            "EndnoteFields" => "Endnotes",
            "TextboxFields" => "Textboxes",
            "HeaderTextboxFields" => "HeaderTextboxes",
            _ => throw new InvalidOperationException("Unsupported field PLC location.")
        };
        var part = structure.Parts.FirstOrDefault(x => x.Name == partName);
        if (part == null) throw new InvalidDataException($"The {partName} field PLC has no document part.");
        var length = checked((int)location.Length.Value);
        if (length < 4 || (length - 4) % 6 != 0)
            throw new InvalidDataException("The field PLC has an invalid size.");
        var count = (length - 4) / 6;
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        var plc = new DocStructureNode("Plcfld", location.Name, location.StreamName,
            location.Offset, length);
        plc.Attributes["fieldCharacterCount"] = count.ToString(CultureInfo.InvariantCulture);
        uint previous = 0;
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            if (cp >= part.Length || i > 0 && cp <= previous)
                throw new InvalidDataException("A field character CP is invalid.");
            var recordOffset = (count + 1) * 4 + i * 2;
            var character = (byte)(bytes[recordOffset] & 0x1F);
            if (character is not (0x13 or 0x14 or 0x15))
                throw new InvalidDataException("A field character has an invalid kind.");
            var fld = new DocStructureNode("Fld", $"FieldCharacter{i}", location.StreamName,
                location.Offset.Value + recordOffset, 2);
            fld.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
            fld.Attributes["cp"] = cp.ToString(CultureInfo.InvariantCulture);
            fld.Attributes["globalCp"] = checked(part.CpStart + cp).ToString(CultureInfo.InvariantCulture);
            fld.Attributes["character"] = $"0x{character:X2}";
            fld.Children.Add(new DocStructureNode("fldch", "FieldCharacterType", location.StreamName,
                location.Offset.Value + recordOffset, 1));
            if (character == 0x13)
            {
                var type = new DocStructureNode("flt", "LastParsedFieldType", location.StreamName,
                    location.Offset.Value + recordOffset + 1, 1);
                type.Attributes["value"] = bytes[recordOffset + 1].ToString(CultureInfo.InvariantCulture);
                fld.Attributes["fieldType"] = type.Attributes["value"];
                fld.Children.Add(type);
            }
            if (character == 0x15)
                fld.Children.Add(new DocStructureNode("grffldEnd", "FieldEndFlags", location.StreamName,
                    location.Offset.Value + recordOffset + 1, 1));
            plc.Children.Add(fld);
            previous = cp;
        }
        location.Children.Add(plc);
    }
}
