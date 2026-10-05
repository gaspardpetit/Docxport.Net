using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes the bounded entries of a standard 16-bit-count STTB.</summary>
internal static class DocStringTableNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var length = checked((int)location.Length.Value);
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        if (length < 4) throw new InvalidDataException("The string table is too short.");
        var extended = U16(bytes, 0) == 0xFFFF;
        var cursor = extended ? 2 : 0;
        var wideCount = location.Name == "ProtectionBookmarkNames";
        var headerBytes = wideCount ? 6 : 4;
        if (length - cursor < headerBytes)
            throw new InvalidDataException("The string table header is truncated.");
        var count = wideCount ? BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(cursor, 4)) :
            U16(bytes, cursor);
        var extraSize = U16(bytes, cursor + (wideCount ? 4 : 2));
        cursor += headerBytes;
        if (count < 0 || count > length - cursor)
            throw new InvalidDataException("The string table has an invalid count.");
        var kind = location.Name switch
        {
            "BookmarkNames" => "SttbfBkmk",
            "FontTable" => "SttbfFfn",
            "AssociatedStrings" => "SttbfAssoc",
            "CommentBookmarkNames" => "SttbfAtnBkmk",
            "RevisionAuthors" => "SttbfRMark",
            "ListNames" => "SttbListNames",
            "FactoidBookmarkNames" => "SttbfBkmkFactoid",
            "ConsistencyBookmarkNames" => "SttbfBkmkFcc",
            "RepairBookmarkNames" => "SttbfBkmkBPRepairs",
            "Captions" => "SttbfCaption",
            "AutoCaptions" => "SttbfAutoCaption",
            "SaveHistory" => "SttbSavedBy",
            "ExternalFileNames" => "SttbFnm",
            "GlossaryNames" => "SttbfGlsy",
            "GlossaryStyles" => "SttbGlsyStyle",
            "ListTemplates" => "SttbRgtplc",
            "ProtectionBookmarkNames" => "SttbfBkmkProt",
            "ProtectionUsers" => "SttbProtUser",
            _ => throw new InvalidOperationException("Unsupported string table location.")
        };
        var expectedExtra = kind switch
        {
            "SttbfCaption" => 6,
            "SttbfAutoCaption" => 2,
            "SttbFnm" => 8,
            "SttbSavedBy" => 0,
            "SttbfGlsy" => 4,
            "SttbGlsyStyle" => 1,
            "SttbRgtplc" => 0,
            "SttbfBkmkProt" => 8,
            "SttbProtUser" => 2,
            "SttbfAtnBkmk" => 10,
            _ => -1
        };
        if (expectedExtra >= 0 && (!extended || extraSize != expectedExtra))
            throw new InvalidDataException($"{kind} has an invalid string-table header.");
        var table = new DocStructureNode(kind, kind, location.StreamName, location.Offset, length);
        table.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        table.Attributes["extended"] = extended ? "true" : "false";
        table.Attributes["extraBytes"] = extraSize.ToString(CultureInfo.InvariantCulture);
        var sttb = new DocStructureNode("STTB", "StringTable", location.StreamName,
            location.Offset, length);
        table.Children.Add(sttb);
        for (var i = 0; i < count; i++)
        {
            var lengthBytes = extended ? 2 : 1;
            if (cursor > length - lengthBytes) throw new InvalidDataException("A string table entry is truncated.");
            var characters = extended ? U16(bytes, cursor) : bytes[cursor];
            if (kind == "SttbfBkmkProt" && characters != 0)
                throw new InvalidDataException("A protection bookmark name must be empty.");
            if (kind == "SttbfAtnBkmk" && characters != 0)
                throw new InvalidDataException("An annotation bookmark name must be empty.");
            if (kind == "SttbRgtplc" && characters is not (0 or 18))
                throw new InvalidDataException("A list template entry has an invalid length.");
            var contentBytes = checked(characters * (extended ? 2 : 1));
            var entryLength = checked(lengthBytes + contentBytes + extraSize);
            if (entryLength > length - cursor)
                throw new InvalidDataException("A string table entry extends beyond its range.");
            var entry = new DocStructureNode("STTBEntry", $"Entry{i}", location.StreamName,
                location.Offset.Value + cursor, entryLength);
            entry.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["characters"] = characters.ToString(CultureInfo.InvariantCulture);
            if (extended && kind is not ("SttbRgtplc" or "SttbfBkmkFactoid" or "SttbfBkmkFcc"))
                entry.Attributes["text"] = System.Text.Encoding.Unicode.GetString(bytes,
                    cursor + lengthBytes, contentBytes);
            var dataKind = kind switch
            {
                "SttbfFfn" => "FFN",
                "SttbfBkmkFactoid" when extended && extraSize == 0 && characters == 6 => "FACTOIDINFO",
                "SttbfBkmkFcc" when extended && extraSize == 0 && characters == 10 => "DPCID",
                _ => "STTBData"
            };
            var data = new DocStructureNode(dataKind, "Data",
                location.StreamName, location.Offset.Value + cursor + lengthBytes, contentBytes);
            entry.Children.Add(data);
            if (dataKind == "FACTOIDINFO")
            {
                data.Attributes["id"] = BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(cursor + lengthBytes)).ToString(CultureInfo.InvariantCulture);
                var fto = new DocStructureNode("FTO", "Recognizer", location.StreamName,
                    data.Offset + 6, 2);
                fto.Attributes["value"] = U16(bytes, cursor + lengthBytes + 6)
                    .ToString(CultureInfo.InvariantCulture);
                data.Children.Add(fto);
            }
            if (dataKind == "DPCID")
            {
                var idpci = new DocStructureNode("IDPCI", "IssueType", location.StreamName,
                    data.Offset + 6, 4);
                idpci.Attributes["value"] = BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(cursor + lengthBytes + 6)).ToString(CultureInfo.InvariantCulture);
                data.Children.Add(idpci);
                var fcct = new DocStructureNode("FCCT", "IssueFlags", location.StreamName,
                    data.Offset + 14, 1);
                fcct.Attributes["value"] = bytes[cursor + lengthBytes + 14]
                    .ToString(CultureInfo.InvariantCulture);
                data.Children.Add(fcct);
            }
            if (kind == "SttbfFfn" && contentBytes >= 39)
            {
                var fontOffset = location.Offset.Value + cursor + lengthBytes;
                var ffid = new DocStructureNode("FFID", "FontFamily", location.StreamName,
                    fontOffset, 1);
                var id = bytes[cursor + lengthBytes];
                ffid.Attributes["pitch"] = (id & 3).ToString(CultureInfo.InvariantCulture);
                ffid.Attributes["trueType"] = (id & 4) != 0 ? "true" : "false";
                ffid.Attributes["family"] = ((id >> 4) & 7).ToString(CultureInfo.InvariantCulture);
                data.Children.Add(ffid);
                data.Children.Add(new DocStructureNode("PANOSE", "FontClassification",
                    location.StreamName, fontOffset + 5, 10));
            }
            if (kind == "SttbRgtplc")
                for (var level = 0; level < contentBytes / 4; level++)
                    data.Children.Add(DocTemplateCodeNavigator.Create(location.StreamName,
                        location.Offset.Value + cursor + lengthBytes + level * 4L,
                        bytes.AsSpan(cursor + lengthBytes + level * 4, 4), $"LevelTemplate{level}"));
            if (extraSize != 0)
            {
                var extraKind = kind switch
                {
                    "SttbfCaption" => "CAPI",
                    "SttbFnm" => "FNIF",
                    "SttbfGlsy" => "LEGOXTR_V11",
                    "SttbfBkmkProt" => "PRTI",
                    "SttbfAtnBkmk" => "ATNBE",
                    _ => "STTBExtra"
                };
                var extra = new DocStructureNode(extraKind, "ExtraData", location.StreamName,
                    location.Offset.Value + cursor + lengthBytes + contentBytes, extraSize);
                if (kind == "SttbfAutoCaption" && extraSize == 2)
                    extra.Attributes["captionIndex"] = U16(bytes, cursor + lengthBytes + contentBytes)
                        .ToString(CultureInfo.InvariantCulture);
                if (kind == "SttbFnm")
                {
                    var extraOffset = cursor + lengthBytes + contentBytes;
                    var fnpi = new DocStructureNode("FNPI", "FileNamePointer", location.StreamName,
                        location.Offset.Value + extraOffset, 2);
                    var pointer = U16(bytes, extraOffset);
                    fnpi.Attributes["type"] = (pointer & 0xF).ToString(CultureInfo.InvariantCulture);
                    fnpi.Attributes["index"] = (pointer >> 4).ToString(CultureInfo.InvariantCulture);
                    extra.Children.Add(fnpi);
                    var fnfb = new DocStructureNode("FNFB", "FileNameFlags", location.StreamName,
                        location.Offset.Value + extraOffset + 3, 1);
                    var flags = bytes[extraOffset + 3];
                    fnfb.Attributes["fat"] = (flags & 1) != 0 ? "true" : "false";
                    fnfb.Attributes["ntfs"] = (flags & 8) != 0 ? "true" : "false";
                    fnfb.Attributes["nonFileSystem"] = (flags & 16) != 0 ? "true" : "false";
                    extra.Children.Add(fnfb);
                }
                if (kind == "SttbfAtnBkmk")
                {
                    var extraOffset = cursor + lengthBytes + contentBytes;
                    extra.Attributes["tag"] = BinaryPrimitives.ReadUInt32LittleEndian(
                        bytes.AsSpan(extraOffset + 2, 4)).ToString(CultureInfo.InvariantCulture);
                }
                if (kind == "SttbfBkmkProt" && extraSize == 8)
                {
                    var fieldOffset = extra.Offset!.Value;
                    var user = new DocStructureNode("UidSel", "PermittedEditors", location.StreamName,
                        fieldOffset, 2);
                    var value = BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(
                        cursor + lengthBytes + contentBytes, 2));
                    user.Attributes["value"] = value.ToString(CultureInfo.InvariantCulture);
                    if (value <= 0)
                        user.Children.Add(new DocStructureNode("UID", "SpecialUser", location.StreamName,
                            fieldOffset, 2));
                    extra.Children.Add(user);
                    extra.Children.Add(new DocStructureNode("ProtectionType", "Protection", location.StreamName,
                        fieldOffset + 2, 2));
                }
                entry.Children.Add(extra);
            }
            sttb.Children.Add(entry);
            cursor += entryLength;
        }
        if (cursor != length)
            throw new InvalidDataException($"{kind} has {length - cursor} trailing bytes.");
        location.Children.Add(table);
    }

    private static ushort U16(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));
}
