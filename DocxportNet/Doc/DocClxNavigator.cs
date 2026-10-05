using System.Buffers.Binary;

namespace DocxportNet.Doc;

/// <summary>Indexes CLX property blocks and text-piece descriptors without reading text bytes.</summary>
internal static class DocClxNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode clx)
    {
        if (clx.Children.Count != 0 || clx.StreamName == null || clx.Offset == null || clx.Length == null)
            return;
        using var table = structure.OpenStream(clx.StreamName);
        using var word = structure.OpenStream("WordDocument");
        var cursor = clx.Offset.Value;
        var end = checked(cursor + clx.Length.Value);
        if (cursor < 0 || end > table.Length || cursor == end)
            throw new InvalidDataException("The CLX range is invalid.");

        while (cursor < end)
        {
            table.Position = cursor;
            var tag = table.ReadByte();
            if (tag == 1)
            {
                Require(cursor, 3, end);
                var size = ReadI16(table);
                if (size < 0 || size > 0x3FA2)
                    throw new InvalidDataException("A CLX property block has an invalid size.");
                var blockLength = 3L + size;
                Require(cursor, blockLength, end);
                var node = new DocStructureNode("Prc", "PropertyBlock", clx.StreamName, cursor, blockLength);
                node.Attributes["propertyBytes"] = size.ToString(System.Globalization.CultureInfo.InvariantCulture);
                var data = new DocStructureNode("PrcData", "PropertyData", clx.StreamName,
                    cursor + 1, size + 2L);
                data.Attributes["propertyBytes"] = size.ToString(System.Globalization.CultureInfo.InvariantCulture);
                node.Children.Add(data);
                clx.Children.Add(node);
                cursor += blockLength;
                continue;
            }
            if (tag != 2)
                throw new InvalidDataException($"Unexpected CLX record type {tag}.");

            Require(cursor, 5, end);
            var plcLength = ReadU32(table);
            if (plcLength < 16 || (plcLength - 4) % 12 != 0)
                throw new InvalidDataException("The CLX piece table has an invalid size.");
            Require(cursor, 5L + plcLength, end);
            if (cursor + 5L + plcLength != end)
                throw new InvalidDataException("The CLX piece table must be its final record.");
            var pcdt = new DocStructureNode("Pcdt", "PieceTableContainer", clx.StreamName,
                cursor, 5L + plcLength);
            var plc = new DocStructureNode("PlcPcd", "PieceTable", clx.StreamName,
                cursor + 5, plcLength);
            var pieceCount = checked((int)((plcLength - 4) / 12));
            plc.Attributes["pieceCount"] = pieceCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            pcdt.Children.Add(plc);
            clx.Children.Add(pcdt);
            ReadPieces(structure, table, word.Length, plc, pieceCount, structure.Parts);
            return;
        }
        throw new InvalidDataException("The CLX has no piece table.");
    }

    private static void ReadPieces(DocStructure structure, Stream table, long wordLength, DocStructureNode plc, int count,
        IReadOnlyList<DocPartRange> parts)
    {
        var cpOffset = plc.Offset!.Value;
        var pcdOffset = cpOffset + (count + 1L) * 4;
        table.Position = cpOffset;
        var previousCp = ReadU32(table);
        if (previousCp != 0) throw new InvalidDataException("The first text CP must be zero.");
        for (var i = 0; i < count; i++)
        {
            table.Position = cpOffset + (i + 1L) * 4;
            var nextCp = ReadU32(table);
            if (nextCp <= previousCp || nextCp >= 0x80000000)
                throw new InvalidDataException("Text-piece CPs must increase within the valid range.");
            var recordOffset = pcdOffset + i * 8L;
            table.Position = recordOffset;
            var flags = ReadU16(table);
            var fc = ReadU32(table);
            var prm = ReadU16(table);
            if ((fc & 0x80000000) != 0)
                throw new InvalidDataException("A text-piece location has a reserved bit set.");
            var compressed = (fc & 0x40000000) != 0;
            var encodedOffset = fc & 0x3FFFFFFF;
            if (compressed && (encodedOffset & 1) != 0)
                throw new InvalidDataException("A compressed text-piece location is not aligned.");
            var textOffset = compressed ? encodedOffset / 2L : encodedOffset;
            var textLength = (long)(nextCp - previousCp) * (compressed ? 1 : 2);
            if (textOffset > wordLength || textLength > wordLength - textOffset)
                throw new InvalidDataException("A text piece points outside WordDocument.");
            var piece = new DocStructureNode("Pcd", $"Piece{i}", plc.StreamName, recordOffset, 8);
            piece.Attributes["cpStart"] = previousCp.ToString(System.Globalization.CultureInfo.InvariantCulture);
            piece.Attributes["cpEnd"] = nextCp.ToString(System.Globalization.CultureInfo.InvariantCulture);
            piece.Attributes["textStream"] = "WordDocument";
            piece.Attributes["textOffset"] = textOffset.ToString(System.Globalization.CultureInfo.InvariantCulture);
            piece.Attributes["textLength"] = textLength.ToString(System.Globalization.CultureInfo.InvariantCulture);
            piece.Attributes["encoding"] = compressed ? "compressed" : "utf16";
            piece.Attributes["flags"] = $"0x{flags:X4}";
            piece.Attributes["prm"] = $"0x{prm:X4}";
            var propertyReference = new DocStructureNode("Prm", "PropertyReference",
                plc.StreamName, recordOffset + 6, 2);
            var complex = (prm & 1) != 0;
            var variant = new DocStructureNode(complex ? "Prm1" : "Prm0", "PropertyReferenceData",
                plc.StreamName, recordOffset + 6, 2);
            if (complex)
                variant.Attributes["propertyBlockIndex"] = (prm >> 1)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
            else
            {
                variant.Attributes["modifierIndex"] = ((prm >> 1) & 0x7F)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
                variant.Attributes["operand"] = (prm >> 8)
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            propertyReference.Children.Add(variant);
            piece.Children.Add(propertyReference);
            var pieceStartCp = previousCp;
            piece.SetPayloadFactory(() => structure.ReadTextPiece(pieceStartCp, nextCp, textOffset,
                checked((int)textLength), compressed));
            AddPartSpans(piece, previousCp, nextCp, parts);
            plc.Children.Add(piece);
            previousCp = nextCp;
        }
        if (parts.Count != 0)
        {
            var expected = checked(parts[parts.Count - 1].CpEnd + (parts.Count > 1 ? 1u : 0u));
            if (previousCp != expected)
                throw new InvalidDataException("The piece table's final CP does not match the document-part counts.");
        }
    }

    private static void AddPartSpans(DocStructureNode piece, uint cpStart, uint cpEnd,
        IReadOnlyList<DocPartRange> parts)
    {
        var spans = new List<(string Name, uint Start, uint End)>();
        var terminalStart = parts.Count == 0 ? 0u : parts[parts.Count - 1].CpEnd;
        var hasTerminal = parts.Count > 1;
        var cursor = cpStart;
        while (cursor < cpEnd)
        {
            DocPartRange? found = null;
            foreach (var part in parts)
                if (part.CpStart <= cursor && cursor < part.CpEnd) { found = part; break; }
            if (found != null)
            {
                var end = Math.Min(cpEnd, found.CpEnd);
                spans.Add((found.Name, cursor, end));
                cursor = end;
            }
            else if (hasTerminal && cursor == terminalStart)
            {
                spans.Add(("TerminalParagraphMark", cursor, cursor + 1));
                cursor++;
            }
            else
            {
                var next = cpEnd;
                foreach (var part in parts)
                    if (part.CpStart > cursor) next = Math.Min(next, part.CpStart);
                if (hasTerminal && terminalStart > cursor) next = Math.Min(next, terminalStart);
                spans.Add(("Unassigned", cursor, next));
                cursor = next;
            }
        }
        if (spans.Count == 1)
        {
            piece.Attributes["part"] = spans[0].Name;
            return;
        }
        foreach (var span in spans)
        {
            var node = new DocStructureNode("PartSpan", span.Name);
            node.Attributes["cpStart"] = span.Start.ToString(System.Globalization.CultureInfo.InvariantCulture);
            node.Attributes["cpEnd"] = span.End.ToString(System.Globalization.CultureInfo.InvariantCulture);
            piece.Children.Add(node);
        }
    }

    private static void Require(long offset, long length, long end)
    {
        if (offset < 0 || length < 0 || offset > end || length > end - offset)
            throw new InvalidDataException("A CLX record extends outside the CLX range.");
    }

    private static ushort ReadU16(Stream stream)
    {
        var bytes = new byte[2];
        ReadFully(stream, bytes);
        return BinaryPrimitives.ReadUInt16LittleEndian(bytes);
    }

    private static short ReadI16(Stream stream) => unchecked((short)ReadU16(stream));

    private static uint ReadU32(Stream stream)
    {
        var bytes = new byte[4];
        ReadFully(stream, bytes);
        return BinaryPrimitives.ReadUInt32LittleEndian(bytes);
    }

    private static void ReadFully(Stream stream, byte[] bytes)
    {
        var read = 0;
        while (read < bytes.Length)
        {
            var count = stream.Read(bytes, read, bytes.Length - read);
            if (count == 0) throw new InvalidDataException("The CLX ended unexpectedly.");
            read += count;
        }
    }
}
