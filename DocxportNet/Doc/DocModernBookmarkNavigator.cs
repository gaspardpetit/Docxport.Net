using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes Word 2003 protection and structured-document-tag bookmark PLCs.</summary>
internal static class DocModernBookmarkNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var sdt = location.Name.StartsWith("Sdt", StringComparison.Ordinal);
        var starts = location.Name.EndsWith("Starts", StringComparison.Ordinal);
        var kind = starts ? sdt ? "Plcbkfd" : "Plcbkf" : sdt ? "Plcbkld" : "Plcbkl";
        var recordKind = starts ? sdt ? "BKFD" : "BKF" : sdt ? "BKLD" : "";
        var recordSize = starts ? sdt ? 10 : 6 : sdt ? 8 : 0;
        var length = checked((int)location.Length.Value);
        if (length < 4 || (length - 4) % (4 + recordSize) != 0)
            throw new InvalidDataException($"{kind} has an invalid length.");
        var count = (length - 4) / (4 + recordSize);
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        var plc = new DocStructureNode(kind, location.Name, location.StreamName,
            location.Offset, length);
        plc.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        var recordsOffset = checked((count + 1) * 4);
        for (var i = 1; i <= count; i++)
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4)) <
                BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i - 1) * 4, 4)))
                throw new InvalidDataException($"{kind} has decreasing CPs.");
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            var entry = new DocStructureNode(kind + "Entry", $"Bookmark{i}", location.StreamName,
                location.Offset.Value + i * 4L, 4);
            entry.Attributes["cp"] = cp.ToString(CultureInfo.InvariantCulture);
            if (recordSize > 0)
            {
                var offset = recordsOffset + i * recordSize;
                var record = new DocStructureNode(recordKind, $"Record{i}", location.StreamName,
                    location.Offset.Value + offset, recordSize);
                record.Attributes[starts ? "endIndex" : "startIndex"] =
                    BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4))
                    .ToString(CultureInfo.InvariantCulture);
                if (starts)
                {
                    var bkf = sdt ? new DocStructureNode("BKF", "BaseBookmarkRecord",
                        location.StreamName, location.Offset.Value + offset, 6) : record;
                    bkf.Children.Add(new DocStructureNode("BKC", "BookmarkFlags",
                        location.StreamName, location.Offset.Value + offset + 4, 2));
                    if (sdt) record.Children.Add(bkf);
                }
                if (sdt)
                {
                    record.Attributes["depth"] = BinaryPrimitives.ReadInt32LittleEndian(
                        bytes.AsSpan(offset + (starts ? 6 : 4), 4)).ToString(CultureInfo.InvariantCulture);
                    if (!starts)
                        record.Children.Add(new DocStructureNode("BKL", "BaseBookmarkRecord",
                            location.StreamName, location.Offset.Value + offset, 4));
                }
                entry.Children.Add(record);
            }
            plc.Children.Add(entry);
        }
        location.Children.Add(plc);
    }
}
