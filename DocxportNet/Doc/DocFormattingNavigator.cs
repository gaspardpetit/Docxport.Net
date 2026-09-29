using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes formatting page tables and lazily expands referenced FKP pages.</summary>
internal static class DocFormattingNavigator
{
    private const int PageSize = 512;

    public static void ExpandPageTable(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length == null) return;
        var length = location.Length.Value;
        if (length < 12 || (length - 4) % 8 != 0)
            throw new InvalidDataException("The formatting page table has an invalid size.");
        var count = checked((int)((length - 4) / 8));
        var start = location.Offset.Value;
        var end = checked(start + length);
        using var table = structure.OpenStream(location.StreamName);
        using var word = structure.OpenStream("WordDocument");
        if (start < 0 || end > table.Length)
            throw new InvalidDataException("The formatting page table points outside its stream.");
        var character = location.Name == "CharacterFormatting";
        var plc = new DocStructureNode(character ? "PlcBteChpx" : "PlcBtePapx",
            "FormattingPageTable", location.StreamName, start, length);
        plc.Attributes["pageCount"] = count.ToString(CultureInfo.InvariantCulture);
        var recordStart = start + (count + 1L) * 4;
        table.Position = start;
        var previous = ReadU32(table);
        for (var i = 0; i < count; i++)
        {
            table.Position = start + (i + 1L) * 4;
            var next = ReadU32(table);
            if (next <= previous)
                throw new InvalidDataException("Formatting page ranges must increase.");
            table.Position = recordStart + i * 4L;
            var pn = ReadU32(table) & 0x003FFFFF;
            var pageOffset = pn * (long)PageSize;
            if (pageOffset > word.Length - PageSize)
                throw new InvalidDataException("A formatting page points outside WordDocument.");
            var bte = new DocStructureNode(character ? "PnFkpChpx" : "PnFkpPapx",
                $"PageReference{i}", location.StreamName, recordStart + i * 4L, 4);
            bte.Attributes["fcStart"] = previous.ToString(CultureInfo.InvariantCulture);
            bte.Attributes["fcEnd"] = next.ToString(CultureInfo.InvariantCulture);
            bte.Attributes["pageNumber"] = pn.ToString(CultureInfo.InvariantCulture);
            var page = new DocStructureNode(character ? "ChpxFkp" : "PapxFkp",
                $"FormattingPage{i}", "WordDocument", pageOffset, PageSize);
            page.Attributes["pageNumber"] = pn.ToString(CultureInfo.InvariantCulture);
            page.Attributes["fcStart"] = previous.ToString(CultureInfo.InvariantCulture);
            page.Attributes["fcEnd"] = next.ToString(CultureInfo.InvariantCulture);
            bte.Children.Add(page);
            plc.Children.Add(bte);
            previous = next;
        }
        location.Children.Add(plc);
    }

    public static void ExpandPage(DocStructure structure, DocStructureNode pageNode)
    {
        if (pageNode.Children.Count != 0 || pageNode.Offset == null) return;
        var page = structure.ReadRange("WordDocument", pageNode.Offset.Value, PageSize);
        var count = page[511];
        var paragraph = pageNode.Kind == "PapxFkp";
        var maxCount = paragraph ? 29 : 101;
        var recordSize = paragraph ? 13 : 1;
        if (count == 0 || count > maxCount || (count + 1) * 4 + count * recordSize > 511)
            throw new InvalidDataException("A formatting page has an invalid run count.");
        pageNode.Attributes[paragraph ? "paragraphCount" : "runCount"] =
            count.ToString(CultureInfo.InvariantCulture);
        var recordStart = (count + 1) * 4;
        var previous = U32(page, 0);
        for (var i = 0; i < count; i++)
        {
            var next = U32(page, (i + 1) * 4);
            if (next <= previous)
                throw new InvalidDataException("Formatting run offsets must increase.");
            var recordOffset = recordStart + i * recordSize;
            var offsetWords = page[recordOffset];
            var run = new DocStructureNode(paragraph ? "PapxRange" : "ChpxRange",
                $"Range{i}", "WordDocument", pageNode.Offset + recordOffset, recordSize);
            run.Attributes["fcStart"] = previous.ToString(CultureInfo.InvariantCulture);
            run.Attributes["fcEnd"] = next.ToString(CultureInfo.InvariantCulture);
            run.Attributes["propertyOffsetWords"] = offsetWords.ToString(CultureInfo.InvariantCulture);
            if (paragraph)
            {
                var bx = new DocStructureNode("BxPap", "ParagraphPropertyReference", "WordDocument",
                    pageNode.Offset + recordOffset, 13);
                bx.Attributes["offsetWords"] = offsetWords.ToString(CultureInfo.InvariantCulture);
                run.Children.Add(bx);
            }
            if (offsetWords != 0)
            {
                var propertyOffset = offsetWords * 2;
                if (propertyOffset < recordStart + count * recordSize || propertyOffset >= 511)
                    throw new InvalidDataException("A formatting property block has an invalid offset.");
                var propertyLength = paragraph
                    ? PapxLength(page, propertyOffset)
                    : 1 + page[propertyOffset];
                if (propertyLength > 511 - propertyOffset)
                    throw new InvalidDataException("A formatting property block extends beyond its page.");
                var property = new DocStructureNode(paragraph ? "PapxInFkp" : "Chpx",
                    "DirectProperties", "WordDocument", pageNode.Offset + propertyOffset, propertyLength);
                property.Attributes["propertyBytes"] = (propertyLength - 1).ToString(CultureInfo.InvariantCulture);
                run.Children.Add(property);
            }
            pageNode.Children.Add(run);
            previous = next;
        }
    }

    private static int PapxLength(byte[] page, int offset)
    {
        var cb = page[offset];
        if (cb != 0) return cb * 2;
        if (offset >= 510) throw new InvalidDataException("A PAPX header is truncated.");
        var extended = page[offset + 1];
        if (extended == 0) throw new InvalidDataException("A PAPX has an invalid extended size.");
        return 2 + extended * 2;
    }

    private static uint U32(byte[] bytes, int offset) =>
        BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset));

    private static uint ReadU32(Stream stream)
    {
        var bytes = new byte[4];
        var read = 0;
        while (read < bytes.Length)
        {
            var count = stream.Read(bytes, read, bytes.Length - read);
            if (count == 0) throw new InvalidDataException("The formatting table ended unexpectedly.");
            read += count;
        }
        return U32(bytes, 0);
    }
}
