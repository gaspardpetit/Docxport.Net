using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Locates style headers, names, and length-prefixed formatting blocks.</summary>
internal static class DocStyleDefinitionNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode definition)
    {
        if (definition.Children.Count != 0 || definition.StreamName == null ||
            definition.Offset == null || definition.Length == null) return;
        var bytes = structure.ReadRange(definition.StreamName, definition.Offset.Value,
            checked((int)definition.Length.Value));
        var baseSize = int.Parse(definition.Attributes["baseSize"], CultureInfo.InvariantCulture);
        if (baseSize is not (10 or 18) || bytes.Length < baseSize + 4)
            throw new InvalidDataException("A style definition has an invalid base size.");
        var stdf = new DocStructureNode("Stdf", "StyleFields", definition.StreamName,
            definition.Offset, baseSize);
        var stdfBase = new DocStructureNode("StdfBase", "BaseFields", definition.StreamName,
            definition.Offset, 10);
        var kind = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2, 2)) & 0xF;
        var propertyCount = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(4, 2)) & 0xF;
        stdfBase.Attributes["styleType"] = kind.ToString(CultureInfo.InvariantCulture);
        stdfBase.Attributes["invariantStyleId"] =
            (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(0, 2)) & 0x0FFF)
            .ToString(CultureInfo.InvariantCulture);
        stdfBase.Attributes["propertySetCount"] = propertyCount.ToString(CultureInfo.InvariantCulture);
        stdfBase.Attributes["basedOnStyleIndex"] =
            (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2, 2)) >> 4)
            .ToString(CultureInfo.InvariantCulture);
        stdfBase.Attributes["nextStyleIndex"] =
            (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(4, 2)) >> 4)
            .ToString(CultureInfo.InvariantCulture);
        stdfBase.Children.Add(new DocStructureNode("GRFSTD", "StyleFlags", definition.StreamName,
            definition.Offset + 8, 2));
        stdf.Children.Add(stdfBase);
        if (baseSize == 18)
        {
            var post = new DocStructureNode("StdfPost2000OrNone", "Post2000Fields",
                definition.StreamName, definition.Offset + 10, 8);
            var postData = new DocStructureNode("StdfPost2000", "Post2000Data",
                definition.StreamName, definition.Offset + 10, 8);
            postData.Attributes["linkedStyleIndex"] =
                (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(10, 2)) & 0x0FFF)
                .ToString(CultureInfo.InvariantCulture);
            post.Children.Add(postData);
            stdf.Children.Add(post);
        }
        definition.Children.Add(stdf);

        var characters = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(baseSize, 2));
        var nameSize = checked(4 + characters * 2);
        if (nameSize > bytes.Length - baseSize ||
            BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(baseSize + nameSize - 2, 2)) != 0)
            throw new InvalidDataException("A style name is invalid.");
        var name = new DocStructureNode("Xstz", "StyleName", definition.StreamName,
            definition.Offset + baseSize, nameSize);
        name.Attributes["text"] = System.Text.Encoding.Unicode.GetString(bytes,
            baseSize + 2, characters * 2);
        name.Children.Add(new DocStructureNode("Xst", "NameCharacters", definition.StreamName,
            definition.Offset + baseSize, nameSize - 2));
        definition.Children.Add(name);

        var cursor = baseSize + nameSize;
        if (cursor == bytes.Length) return;
        var group = new DocStructureNode("GrLPUpxSw", "StyleProperties", definition.StreamName,
            definition.Offset + cursor, bytes.Length - cursor);
        var groupKind = kind switch
        {
            1 => "StkParaGRLPUPX",
            2 => "StkCharGRLPUPX",
            3 => "StkTableGRLPUPX",
            4 => "StkListGRLPUPX",
            _ => "UnknownStyleProperties"
        };
        var typed = new DocStructureNode(groupKind, "TypedStyleProperties", definition.StreamName,
            definition.Offset + cursor, bytes.Length - cursor);
        group.Children.Add(typed);
        definition.Children.Add(group);
        if (groupKind == "UnknownStyleProperties") return;
        var types = kind switch
        {
            1 => new[] { "Papx", "Chpx" },
            2 => new[] { "Chpx" },
            3 => new[] { "Tapx", "Papx", "Chpx" },
            _ => new[] { "Papx" }
        };
        for (var i = 0; i < Math.Min(propertyCount, types.Length); i++)
        {
            if (bytes.Length - cursor < 2) throw new InvalidDataException("A style property block is truncated.");
            var contentLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor, 2));
            var total = checked(2 + contentLength + (contentLength & 1));
            if (total > bytes.Length - cursor)
                throw new InvalidDataException("A style property block extends beyond the style.");
            var type = types[i];
            var block = new DocStructureNode("LPUpx" + type, $"PropertySet{i}",
                definition.StreamName, definition.Offset + cursor, total);
            block.Attributes["propertyBytes"] = contentLength.ToString(CultureInfo.InvariantCulture);
            block.Children.Add(new DocStructureNode("Upx" + type, "Properties", definition.StreamName,
                definition.Offset + cursor + 2, contentLength));
            if ((contentLength & 1) != 0)
                block.Children.Add(new DocStructureNode("UPXPadding", "Padding",
                    definition.StreamName, definition.Offset + cursor + 2 + contentLength, 1));
            typed.Children.Add(block);
            cursor += total;
        }
        if (kind is 1 or 2 && propertyCount > types.Length && cursor < bytes.Length)
            AddRevisionProperties(definition, typed, bytes, cursor, kind);
    }

    private static void AddRevisionProperties(DocStructureNode definition,
        DocStructureNode parent, byte[] bytes, int cursor, int styleKind)
    {
        if (bytes.Length - cursor < 10) return;
        var contentLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
        if (contentLength < 8 || contentLength > bytes.Length - cursor - 2) return;
        var prefixKind = styleKind == 1 ? "StkParaLPUpxGrLPUpxRM" : "StkCharLPUpxGrLPUpxRM";
        var bodyKind = styleKind == 1 ? "StkParaUpxGrLPUpxRM" : "StkCharUpxGrLPUpxRM";
        var prefix = new DocStructureNode(prefixKind, "RevisionProperties", definition.StreamName,
            definition.Offset + cursor, 2 + contentLength);
        var body = new DocStructureNode(bodyKind, "RevisionStyle", definition.StreamName,
            definition.Offset + cursor + 2, contentLength);
        if (BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 2)) != 6) return;
        var mark = new DocStructureNode("LPUpxRm", "RevisionMark", definition.StreamName,
            definition.Offset + cursor + 2, 8);
        var details = new DocStructureNode("UpxRm", "RevisionDetails", definition.StreamName,
            definition.Offset + cursor + 4, 6);
        details.Children.Add(new DocStructureNode("DTTM", "RevisedAt", definition.StreamName,
            details.Offset, 4));
        details.Attributes["authorIndex"] = BinaryPrimitives.ReadInt16LittleEndian(
            bytes.AsSpan(cursor + 8)).ToString(CultureInfo.InvariantCulture);
        mark.Children.Add(details);
        body.Children.Add(mark);
        var inner = cursor + 10;
        var end = cursor + 2 + contentLength;
        var types = styleKind == 1 ? new[] { "Papx", "Chpx" } : new[] { "Chpx" };
        foreach (var type in types)
        {
            if (end - inner < 2) break;
            var payload = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(inner));
            var total = 2 + payload + (payload & 1);
            if (total > end - inner) break;
            var block = new DocStructureNode("LPUpx" + type + "RM", type + "Revision",
                definition.StreamName, definition.Offset + inner, total);
            block.Children.Add(new DocStructureNode("Upx" + type, "Properties",
                definition.StreamName, definition.Offset + inner + 2, payload));
            body.Children.Add(block);
            inner += total;
        }
        prefix.Children.Add(body);
        parent.Children.Add(prefix);
    }
}
