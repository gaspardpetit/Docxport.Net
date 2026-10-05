using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes the column and cell records inside a table-definition modifier.</summary>
internal static class DocTableOperandNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode operand)
    {
        if (operand.Children.Count != 0 || operand.StreamName == null ||
            operand.Offset == null || operand.Length == null) return;
        var bytes = structure.ReadRange(operand.StreamName, operand.Offset.Value,
            checked((int)operand.Length.Value));
        if (bytes.Length < 3) return;
        var count = bytes[2];
        if (count > 63 || 3 + 2 * (count + 1) > bytes.Length) return;
        var declared = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        if (declared + 1 != bytes.Length) return;
        operand.Attributes["columnCount"] = count.ToString(CultureInfo.InvariantCulture);
        var cursor = 3;
        for (var i = 0; i <= count; i++)
        {
            var edge = new DocStructureNode("XAS", $"ColumnEdge{i}", operand.StreamName,
                operand.Offset + cursor, 2);
            edge.Attributes["twips"] = BinaryPrimitives.ReadInt16LittleEndian(bytes.AsSpan(cursor, 2))
                .ToString(CultureInfo.InvariantCulture);
            operand.Children.Add(edge);
            cursor += 2;
        }
        for (var i = 0; cursor + 20 <= bytes.Length; i++, cursor += 20)
        {
            var cell = new DocStructureNode("TC80", $"Cell{i}", operand.StreamName,
                operand.Offset + cursor, 20);
            cell.Children.Add(new DocStructureNode("TCGRF", "CellFlags", operand.StreamName,
                operand.Offset + cursor, 2));
            cell.Attributes["preferredWidth"] = BinaryPrimitives.ReadUInt16LittleEndian(
                bytes.AsSpan(cursor + 2, 2)).ToString(CultureInfo.InvariantCulture);
            var borders = new[] { "Top", "Left", "Bottom", "Right" };
            for (var side = 0; side < borders.Length; side++)
                cell.Children.Add(new DocStructureNode("Brc80MayBeNil", borders[side],
                    operand.StreamName, operand.Offset + cursor + 4 + side * 4L, 4));
            operand.Children.Add(cell);
        }
    }
}
