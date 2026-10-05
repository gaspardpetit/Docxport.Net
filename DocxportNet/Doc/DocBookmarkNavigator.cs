using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes parallel bookmark starts, ends, and their FBKF references.</summary>
internal static class DocBookmarkNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var starts = location.Name.EndsWith("Starts", StringComparison.Ordinal);
        var extended = location.Name.StartsWith("Factoid", StringComparison.Ordinal) ||
            location.Name.StartsWith("Consistency", StringComparison.Ordinal);
        var length = checked((int)location.Length.Value);
        var recordSize = extended ? (starts ? 6 : 4) : starts ? 4 : 0;
        if (length < 4 || (length - 4) % (4 + recordSize) != 0)
            throw new InvalidDataException("The bookmark PLC has an invalid size.");
        var count = (length - 4) / (4 + recordSize);
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        var kind = extended ? (starts ? "Plcfbkfd" : "Plcfbkld") :
            starts ? "Plcfbkf" : "Plcfbkl";
        var plc = new DocStructureNode(kind, location.Name, location.StreamName,
            location.Offset, length);
        plc.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        uint previous = 0;
        var documentEnd = structure.Parts.LastOrDefault()?.CpEnd;
        for (var i = 0; i < count; i++)
        {
            var cp = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            if (i > 0 && cp < previous || documentEnd != null && cp > documentEnd)
                throw new InvalidDataException("A bookmark CP is invalid.");
            var entry = new DocStructureNode(kind + "Entry", $"Bookmark{i}", location.StreamName,
                location.Offset.Value + i * 4L, 4);
            entry.Attributes["index"] = i.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["cp"] = cp.ToString(CultureInfo.InvariantCulture);
            if (recordSize != 0)
            {
                var recordOffset = (count + 1) * 4 + i * recordSize;
                var pairedIndex = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(recordOffset, 2));
                entry.Attributes[starts ? "endIndex" : "startIndex"] =
                    pairedIndex.ToString(CultureInfo.InvariantCulture);
                var record = new DocStructureNode(starts ? extended ? "FBKFD" : "FBKF" : "FBKLD",
                    starts ? "BookmarkStartRecord" : "BookmarkEndRecord", location.StreamName,
                    location.Offset.Value + recordOffset, recordSize);
                if (starts)
                {
                    var fbkf = extended ? new DocStructureNode("FBKF", "BaseBookmarkRecord",
                        location.StreamName, location.Offset.Value + recordOffset, 4) : record;
                    fbkf.Children.Add(new DocStructureNode("BKC", "BookmarkFlags", location.StreamName,
                        location.Offset.Value + recordOffset + 2, 2));
                    if (extended) record.Children.Add(fbkf);
                }
                if (extended)
                    record.Attributes["depth"] = BinaryPrimitives.ReadUInt16LittleEndian(
                        bytes.AsSpan(recordOffset + (starts ? 4 : 2), 2)).ToString(CultureInfo.InvariantCulture);
                entry.Children.Add(record);
            }
            plc.Children.Add(entry);
            previous = cp;
        }
        location.Children.Add(plc);
    }
}
