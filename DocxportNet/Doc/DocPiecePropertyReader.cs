using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Reads the property modifiers attached to a text piece's PCD.</summary>
internal sealed class DocPiecePropertyReader
{
    private readonly DocTextIndex index;
    private readonly DocStructureNode[] propertyBlocks;

    public DocPiecePropertyReader(DocTextIndex index)
    {
        this.index = index;
        propertyBlocks = Descendants(index.Structure.Root)
            .Where(x => x.Kind == "Prc").ToArray();
    }

    public byte[] Read(DocIndexedTextPiece piece, bool character)
    {
        var reference = piece.Node.Children.FirstOrDefault(x => x.Kind == "Prm")?
            .Children.FirstOrDefault();
        if (reference?.Kind == "Prm1")
        {
            var blockIndex = int.Parse(reference.Attributes["propertyBlockIndex"],
                CultureInfo.InvariantCulture);
            if (blockIndex < 0 || blockIndex >= propertyBlocks.Length)
                throw new InvalidDataException("A text piece refers to an invalid CLX property block.");
            var block = propertyBlocks[blockIndex];
            var size = int.Parse(block.Attributes["propertyBytes"],
                CultureInfo.InvariantCulture);
            return Extract(index.Structure.ReadRange(block.StreamName!,
                block.Offset!.Value + 3, size), character ? 2 : 1);
        }
        if (reference?.Kind != "Prm0") return [];
        var modifier = byte.Parse(reference.Attributes["modifierIndex"],
            CultureInfo.InvariantCulture);
        var operand = byte.Parse(reference.Attributes["operand"],
            CultureInfo.InvariantCulture);
        ushort sprm = checked((ushort)(character ? modifier switch
        {
            0x4D => 0x2A0C, // highlight
            0x53 => 0x2A33, // reset to paragraph style
            0x55 => 0x0835, // bold
            0x56 => 0x0836, // italic
            0x57 => 0x0837, // strike
            0x5A => 0x083A, // small caps
            0x5B => 0x083B, // caps
            0x5C => 0x083C, // hidden
            0x5E => 0x2A3E, // underline
            0x62 => 0x2A42, // indexed text color
            0x68 => 0x2A48, // subscript/superscript
            0x75 => 0x0855, // special character
            _ => 0
        } : modifier switch
        {
            0x05 => 0x2461, // justification
            0x07 => 0x2405, // keep lines
            0x08 => 0x2406, // keep with next
            0x09 => 0x2407, // page break before
            0x0C => 0x260A, // list level
            0x18 => 0x2416, // in table
            0x19 => 0x2417, // row terminator
            0x33 => 0x2431, // widow/orphan control
            _ => 0
        }));
        return sprm == 0 ? [] : [(byte)sprm, (byte)(sprm >> 8), operand];
    }

    private static byte[] Extract(ReadOnlySpan<byte> sprms, int category)
    {
        using var output = new MemoryStream();
        for (var offset = 0; offset < sprms.Length;)
        {
            if (offset + 2 > sprms.Length)
                throw new InvalidDataException("A CLX property block has a truncated SPRM.");
            var sprm = BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(offset));
            var operandStart = offset + 2;
            var operandLength = (sprm >> 13) switch
            {
                0 or 1 => 1,
                2 or 4 or 5 => 2,
                3 => 4,
                6 when sprm == 0xD608 && operandStart + 2 <= sprms.Length =>
                    1 + BinaryPrimitives.ReadUInt16LittleEndian(sprms.Slice(operandStart)),
                6 when sprm == 0xC615 && operandStart < sprms.Length &&
                    sprms[operandStart] == 255 =>
                    DocParagraphFormatting.GetExtendedTabOperandLength(
                        sprms.Slice(operandStart)),
                6 when operandStart < sprms.Length => 1 + sprms[operandStart],
                7 => 3,
                _ => throw new InvalidDataException(
                    $"A CLX property block has an invalid SPRM at {offset}: 0x{sprm:X4}.")
            };
            var end = operandStart + operandLength;
            if (end > sprms.Length)
                throw new InvalidDataException("A CLX property block has a truncated operand.");
            if (((sprm >> 10) & 7) == category)
            {
                var bytes = sprms.Slice(offset, end - offset).ToArray();
                output.Write(bytes, 0, bytes.Length);
            }
            offset = end;
        }
        return output.ToArray();
    }

    private static IEnumerable<DocStructureNode> Descendants(DocStructureNode node)
    {
        foreach (var child in node.Children)
        {
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }
}
