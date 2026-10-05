using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Locates later FIB metadata, including comment trees and revision IDs.</summary>
internal static class DocMetadataNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length is not > 0) return;
        var kind = location.Name switch
        {
            "CommentTree" => "AtrdExtra",
            "RevisionSaveIds" => "PLRSID",
            "SmartTagData" => "SmartTagData",
            "PageProperties" => "PGPArray",
            "UserInfoRanges" => "Plcfuim",
            "UserInfoGuids" => "PlfguidUim",
            "CommandCustomizations" => "Tcg",
            "PrinterDriver" => "PrDrvr",
            "PrinterPort" => "PrEnvPort",
            "PrinterEnvironment" => "PrEnvLand",
            "PrintMergeState" => "Pms",
            "UserStrings" => "StwUser",
            "ToolbarData" => "SttbTtmbd",
            "RouteSlip" => "RouteSlip",
            "ThreadingMetadata" => "RmdThreading",
            "OutlineStyles" => "PlfCosl",
            "GrammarOptionSets" => "PlfGosl",
            "OleControlInfo" => "RgxOcxInfo",
            "GrammarCookieData" => "RgCdb",
            "OldCookieRanges" => "PlcfcookieOld",
            "CookieRanges" => "Plcfcookie",
            _ => "OfficeDataSource"
        };
        var node = new DocStructureNode(kind, location.Name, location.StreamName,
            location.Offset, location.Length);
        if (kind == "AtrdExtra" && location.Length.Value % 18 == 0)
        {
            var count = location.Length.Value / 18;
            node.Attributes["commentCount"] = count.ToString(CultureInfo.InvariantCulture);
            for (var i = 0L; i < count; i++)
            {
                var record = new DocStructureNode("ATRDPost10", $"Comment{i}",
                    location.StreamName, location.Offset + i * 18, 18);
                record.Children.Add(new DocStructureNode("DTTM", "ModifiedAt",
                    location.StreamName, record.Offset, 4));
                node.Children.Add(record);
            }
        }
        else if (kind == "PLRSID" && location.Length.Value >= 24)
        {
            var header = structure.ReadRange(location.StreamName, location.Offset.Value, 24);
            var count = BinaryPrimitives.ReadUInt32LittleEndian(header);
            var itemSize = BinaryPrimitives.ReadUInt32LittleEndian(header.AsSpan(4));
            if (itemSize == 4 && count <= (location.Length.Value - 24) / 4 &&
                24 + count * 4L == location.Length.Value)
                node.Attributes["revisionCount"] = count.ToString(CultureInfo.InvariantCulture);
        }
        else if (kind == "PGPArray") ExpandPageProperties(structure, node);
        else if (kind == "Plcfuim") ExpandUserInfo(structure, node);
        else if (kind == "PlfguidUim") ExpandUserGuids(structure, node);
        else if (kind == "PlfGosl") ExpandGrammarOptions(structure, node);
        else if (kind == "PlfCosl") ExpandLegacyGrammarOptions(structure, node);
        else if (kind == "RgxOcxInfo") ExpandOleControls(structure, node);
        else if (kind == "RgCdb") ExpandGrammarCookies(structure, node);
        else if (kind == "SttbTtmbd") ExpandEmbeddedFonts(structure, node);
        else if (kind == "Pms") ExpandPrintMerge(structure, node);
        else if (kind == "RouteSlip") ExpandRouteSlip(structure, node);
        else if (kind is "Plcfcookie" or "PlcfcookieOld") ExpandCookiePlc(structure, node);
        else if (kind == "RmdThreading") ExpandThreading(structure, node);
        location.Children.Add(node);
    }

    private static void ExpandCookiePlc(DocStructure structure, DocStructureNode node)
    {
        var size = node.Kind == "Plcfcookie" ? 10 : 16;
        if (node.Length < 4 || (node.Length - 4) % (4 + size) != 0) return;
        var count = checked((int)((node.Length.Value - 4) / (4 + size)));
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length.Value));
        for (var i = 1; i <= count; i++)
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4)) <
                BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i - 1) * 4))) return;
        var records = 4L * (count + 1);
        for (var i = 0; i < count; i++)
            node.Children.Add(new DocStructureNode(size == 10 ? "FCKS" : "FCKSOLD",
                $"Cookie{i}", node.StreamName, node.Offset + records + i * (long)size, size));
        node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
    }

    private static void ExpandThreading(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 6) return;
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length.Value));
        if (BinaryPrimitives.ReadUInt16LittleEndian(bytes) != 0xFFFF) return;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2));
        var extra = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(4));
        if (extra != 8 || count > (bytes.Length - 6) / 10) return;
        var cursor = 6;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 2) return;
            var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
            var size = 2L + chars * 2L + 8;
            if (size > bytes.Length - cursor) return;
            cursor += (int)size;
        }
        var messages = new DocStructureNode("STTB", "MessageIdentifiers", node.StreamName,
            node.Offset, cursor);
        cursor = 6;
        for (var i = 0; i < count; i++)
        {
            var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
            messages.Children.Add(new DocStructureNode("MDP", $"MessageProperties{i}",
                node.StreamName, node.Offset + cursor + 2 + chars * 2L, 8));
            cursor += 2 + chars * 2 + 8;
        }
        node.Children.Add(messages);
    }

    private static void ExpandPageProperties(DocStructure structure, DocStructureNode node)
    {
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value, checked((int)node.Length!.Value));
        if (bytes.Length < 2) return;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        var cursor = 2;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 14) return;
            var flags = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 12));
            var optionBytes = 0;
            for (var bit = 0; bit <= 8; bit++)
                if ((flags & (1 << bit)) != 0) optionBytes += bit < 4 ? 4 : bit < 8 ? 8 : 2;
            var optionLength = flags == 0 ? 0 : 2 + optionBytes;
            if (bytes.Length - cursor < 14 + optionLength) return;
            if (optionLength != 0 &&
                BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 14)) != optionBytes) return;
            var info = new DocStructureNode("PGPInfo", $"PageProperty{i}", node.StreamName,
                node.Offset + cursor, 14 + optionLength);
            info.Attributes["flags"] = $"0x{flags:X4}";
            if (optionLength != 0)
            {
                var options = new DocStructureNode("PGPOptions", "Options", node.StreamName,
                    node.Offset + cursor + 14, optionLength);
                var field = cursor + 16;
                var names = new[] { "LeftMargin", "RightMargin", "BeforeMargin", "AfterMargin",
                    "LeftBorder", "RightBorder", "TopBorder", "BottomBorder", "Type" };
                for (var bit = 0; bit <= 8; bit++)
                {
                    if ((flags & (1 << bit)) == 0) continue;
                    var size = bit < 4 ? 4 : bit < 8 ? 8 : 2;
                    if (bit >= 4 && bit < 8)
                        options.Children.Add(new DocStructureNode("Brc", names[bit],
                            node.StreamName, node.Offset + field, size));
                    field += size;
                }
                info.Children.Add(options);
            }
            node.Children.Add(info);
            cursor += 14 + optionLength;
        }
        if (cursor == bytes.Length)
            node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
    }

    private static void ExpandUserInfo(DocStructure structure, DocStructureNode node)
    {
        var length = node.Length!.Value;
        if (length < 4 || (length - 4) % 24 != 0) return;
        var count = checked((int)((length - 4) / 24));
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value, checked((int)length));
        node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        var records = 4L * (count + 1);
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4));
            var uim = new DocStructureNode("UIM", $"UserInfo{i}", node.StreamName,
                node.Offset + records + i * 20L, 20);
            uim.Attributes["cp"] = cp.ToString(CultureInfo.InvariantCulture);
            node.Children.Add(uim);
        }
    }

    private static void ExpandUserGuids(DocStructure structure, DocStructureNode node)
    {
        var length = node.Length!.Value;
        if (length < 4) return;
        var count = BinaryPrimitives.ReadUInt32LittleEndian(
            structure.ReadRange(node.StreamName!, node.Offset!.Value, 4));
        if (count <= (length - 4) / 16 && length == 4 + count * 16L)
            node.Attributes["guidCount"] = count.ToString(CultureInfo.InvariantCulture);
    }

    private static void ExpandGrammarOptions(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 4 || (node.Length - 4) % 8 != 0) return;
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value, 4);
        var count = BinaryPrimitives.ReadInt32LittleEndian(bytes);
        if (count < 0 || 4L + count * 8 != node.Length) return;
        node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
        {
            var entry = new DocStructureNode("GOSL", $"OptionSet{i}", node.StreamName,
                node.Offset + 4 + i * 8L, 8);
            entry.Children.Add(new DocStructureNode("LID", "Language", node.StreamName,
                entry.Offset + 2, 2));
            node.Children.Add(entry);
        }
    }

    private static void ExpandLegacyGrammarOptions(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 4 || (node.Length - 4) % 10 != 0) return;
        var count = BinaryPrimitives.ReadInt32LittleEndian(
            structure.ReadRange(node.StreamName!, node.Offset!.Value, 4));
        if (count < 0 || 4L + count * 10 != node.Length) return;
        node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
        {
            var entry = new DocStructureNode("COSL", $"OptionSet{i}", node.StreamName,
                node.Offset + 4 + i * 10L, 10);
            entry.Children.Add(new DocStructureNode("LID", "Language", node.StreamName,
                entry.Offset + 2, 2));
            node.Children.Add(entry);
        }
    }

    private static void ExpandEmbeddedFonts(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 10) return;
        var header = structure.ReadRange(node.StreamName!, node.Offset!.Value, 10);
        var count = BinaryPrimitives.ReadInt16LittleEndian(header.AsSpan(2));
        var capacity = BinaryPrimitives.ReadInt16LittleEndian(header.AsSpan(4));
        var dataOffset = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(8));
        if (count < 0 || count > 64 || capacity != 64 || dataOffset < 10 ||
            dataOffset + count * 12L > node.Length) return;
        var sttb = new DocStructureNode("SttbW6", "EmbeddedFontHeader", node.StreamName,
            node.Offset, 10);
        sttb.Attributes["fontCount"] = count.ToString(CultureInfo.InvariantCulture);
        node.Children.Add(sttb);
        for (var i = 0; i < count; i++)
        {
            var font = new DocStructureNode("Ttmbd", $"Font{i}", node.StreamName,
                node.Offset + dataOffset + i * 12L, 12);
            node.Children.Add(font);
        }
    }

    private static void ExpandOleControls(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 4 || (node.Length - 4) % 20 != 0) return;
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length.Value));
        var count = BinaryPrimitives.ReadUInt32LittleEndian(bytes);
        if (4L + count * 20 != node.Length) return;
        node.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
        {
            var entry = new DocStructureNode("OcxInfo", $"Control{i}", node.StreamName,
                node.Offset + 4 + i * 20L, 20);
            var offset = 4 + i * 20;
            entry.Attributes["cookie"] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset))
                .ToString(CultureInfo.InvariantCulture);
            entry.Attributes["fieldIndex"] = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset + 4))
                .ToString(CultureInfo.InvariantCulture);
            node.Children.Add(entry);
        }
    }

    private static void ExpandGrammarCookies(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 8) return;
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length.Value));
        var total = BinaryPrimitives.ReadUInt32LittleEndian(bytes);
        var count = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(4));
        if (total != bytes.Length || count > (total - 8) / 4) return;
        var cursor = 8;
        var records = new List<DocStructureNode>();
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 4) return;
            var length = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(cursor));
            if (length > bytes.Length - cursor - 4) return;
            records.Add(new DocStructureNode("CDB", $"Cookie{i}", node.StreamName,
                node.Offset + cursor, 4 + length));
            cursor += checked((int)length + 4);
        }
        if (cursor != bytes.Length) return;
        node.Attributes["cookieCount"] = count.ToString(CultureInfo.InvariantCulture);
        foreach (var record in records) node.Children.Add(record);
    }

    private static void ExpandPrintMerge(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 30) return;
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length.Value));
        var sqlLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(28));
        if (sqlLength % 2 != 0 || sqlLength > 512 || 30L + sqlLength > bytes.Length) return;
        node.Children.Add(new DocStructureNode("Wpms", "MergeState", node.StreamName, node.Offset, 2));
        for (var i = 0; i < 2; i++)
        {
            var pmfs = new DocStructureNode("Pmfs", $"DataSource{i}", node.StreamName,
                node.Offset + 8 + i * 8, 8);
            pmfs.Children.Add(new DocStructureNode("FNPI", "FileName", node.StreamName,
                pmfs.Offset + 6, 2));
            node.Children.Add(pmfs);
        }
        var rfs = new DocStructureNode("Rfs", "RecordFilter", node.StreamName, node.Offset + 24, 4);
        rfs.Attributes["stringTableHandle"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(26))
            .ToString(CultureInfo.InvariantCulture);
        node.Children.Add(rfs);
        var cursor = 30 + sqlLength;
        var hasStringTable = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(26)) != 0;
        var parsedStringTable = false;
        if (hasStringTable &&
            bytes.Length - cursor >= 6)
        {
            var sttbLength = StringTableLength(bytes.AsSpan(cursor));
            if (sttbLength > 0)
            {
                node.Children.Add(new DocStructureNode("SttbfRfs", "MergeStrings", node.StreamName,
                    node.Offset + cursor, sttbLength));
                cursor += sttbLength;
                parsedStringTable = true;
            }
        }
        if ((!hasStringTable || parsedStringTable) && bytes.Length - cursor == 4)
            node.Children.Add(new DocStructureNode("Wpmsdt", "DocumentType", node.StreamName,
                node.Offset + cursor, 4));
    }

    private static int StringTableLength(ReadOnlySpan<byte> bytes)
    {
        if (bytes.Length < 6 || BinaryPrimitives.ReadUInt16LittleEndian(bytes) != 0xFFFF)
            return -1;
        var count = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(2));
        var extra = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(4));
        var cursor = 6;
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 2) return -1;
            var chars = BinaryPrimitives.ReadUInt16LittleEndian(bytes.Slice(cursor));
            var length = 2L + chars * 2L + extra;
            if (length > bytes.Length - cursor) return -1;
            cursor += (int)length;
        }
        return cursor;
    }

    private static void ExpandRouteSlip(DocStructure structure, DocStructureNode node)
    {
        if (node.Length < 24) return;
        var bytes = structure.ReadRange(node.StreamName!, node.Offset!.Value,
            checked((int)node.Length.Value));
        var cursor = 16;
        for (var i = 0; i < 4; i++)
        {
            if (bytes.Length - cursor < 2) return;
            var length = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
            if (length > 255 || bytes.Length - cursor - 2 < length) return;
            cursor += 2 + length;
        }
        var count = BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(14));
        if (count < 0) return;
        var entries = new List<DocStructureNode>();
        for (var i = 0; i < count; i++)
        {
            if (bytes.Length - cursor < 4) return;
            var idLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor));
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cursor + 2));
            var length = 4L + idLength + nameLength;
            if (length > bytes.Length - cursor) return;
            entries.Add(new DocStructureNode("RouteSlipInfo", $"Recipient{i}", node.StreamName,
                node.Offset + cursor, length));
            cursor += (int)length;
        }
        if (cursor != bytes.Length) return;
        var protection = new DocStructureNode("RouteSlipProtectionEnum", "Protection",
            node.StreamName, node.Offset + 8, 2);
        protection.Attributes["value"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(8))
            .ToString(CultureInfo.InvariantCulture);
        node.Children.Add(protection);
        foreach (var entry in entries) node.Children.Add(entry);
        node.Attributes["recipientCount"] = count.ToString(CultureInfo.InvariantCulture);
    }
}
