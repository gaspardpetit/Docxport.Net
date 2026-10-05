using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes fixed-record PLCs for document-wide text ranges.</summary>
internal static class DocRangePlcNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var (kind, recordKind, recordSize, allowDuplicates) = location.Name switch
        {
            "SpellingRanges" => ("Plcfspl", "SpellingSpls", 2, true),
            "GrammarRanges" => ("Plcfgram", "GrammarSpls", 2, true),
            "LanguageDetectionRanges" => ("Plcflad", "LadSpls", 2, true),
            "SmartTagRanges" => ("Plcffactoid", "FactoidSpls", 2, true),
            "AutoSummaryRanges" => ("PlcfAsumy", "ASUMY", 4, false),
            "SubdocumentRanges" => ("PlcfWKB", "WKB", 12, false),
            _ => throw new InvalidOperationException("Unsupported range PLC.")
        };
        var length = checked((int)location.Length.Value);
        if (length < 4 || (length - 4) % (4 + recordSize) != 0)
            throw new InvalidDataException($"{kind} has an invalid length.");
        var count = (length - 4) / (4 + recordSize);
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, length);
        var cpBytes = checked((count + 1) * 4);
        var table = new DocStructureNode(kind, location.Name, location.StreamName,
            location.Offset, length);
        table.Attributes["entryCount"] = count.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < count; i++)
        {
            var start = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(i * 4, 4));
            var end = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan((i + 1) * 4, 4));
            if (end < start || !allowDuplicates && end == start)
                throw new InvalidDataException($"{kind} has invalid CP boundaries.");
            var entry = new DocStructureNode(kind + "Entry", $"Range{i}", location.StreamName,
                location.Offset.Value + i * 4L, 4);
            entry.Attributes["cpStart"] = start.ToString(CultureInfo.InvariantCulture);
            entry.Attributes["cpEnd"] = end.ToString(CultureInfo.InvariantCulture);
            var recordOffset = cpBytes + i * recordSize;
            var record = new DocStructureNode(recordKind, $"Record{i}", location.StreamName,
                location.Offset.Value + recordOffset, recordSize);
            if (recordSize == 2)
            {
                record.Attributes["flags"] = $"0x{BinaryPrimitives.ReadUInt16LittleEndian(
                    bytes.AsSpan(recordOffset, 2)):X4}";
                record.Children.Add(new DocStructureNode("SPLS", "State", location.StreamName,
                    record.Offset, 2));
            }
            else if (recordKind == "ASUMY")
                record.Attributes["level"] = BinaryPrimitives.ReadInt32LittleEndian(
                    bytes.AsSpan(recordOffset, 4)).ToString(CultureInfo.InvariantCulture);
            entry.Children.Add(record);
            table.Children.Add(entry);
        }
        location.Children.Add(table);
    }
}
