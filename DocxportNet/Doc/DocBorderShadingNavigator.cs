using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes nested border and shading values inside table operands.</summary>
internal static class DocBorderShadingNavigator
{
    private static readonly string[] Sides =
        ["Top", "Left", "Bottom", "Right", "HorizontalInside", "VerticalInside"];

    public static void Expand(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length == null) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value,
            checked((int)node.Length.Value));
        switch (node.Kind)
        {
            case "TableBordersOperand" when bytes.Length == 49 && bytes[0] == 48:
                AddBorders(node, 1, 8, "Brc");
                break;
            case "TableBordersOperand80" when bytes.Length == 25 && bytes[0] == 24:
                AddBorders(node, 1, 4, "Brc80MayBeNil");
                break;
            case "TableBrcOperand" when bytes.Length == 12 && bytes[0] == 11:
                node.Children.Add(new DocStructureNode("ItcFirstLim", "CellRange",
                    node.StreamName, node.Offset + 1, 2));
                node.Children.Add(new DocStructureNode("BrcMayBeNil", "Border",
                    node.StreamName, node.Offset + 4, 8));
                break;
            case "BrcOperand" when bytes.Length == 9 && bytes[0] == 8:
                node.Children.Add(new DocStructureNode("Brc", "Border", node.StreamName,
                    node.Offset + 1, 8));
                break;
            case "TableShadeOperand" when bytes.Length == 13 && bytes[0] == 12:
                node.Children.Add(new DocStructureNode("ItcFirstLim", "CellRange",
                    node.StreamName, node.Offset + 1, 2));
                node.Children.Add(new DocStructureNode("Shd", "Shading",
                    node.StreamName, node.Offset + 3, 10));
                break;
            case "DefTableShdOperand" when bytes.Length > 1 && (bytes.Length - 1) % 10 == 0:
                AddShadingArray(node, 10, "Shd");
                break;
            case "DefTableShd80Operand" when bytes.Length > 1 && (bytes.Length - 1) % 2 == 0:
                AddShadingArray(node, 2, "Shd80");
                break;
            case "BrcMayBeNil" when bytes.Length == 8:
                node.Children.Add(new DocStructureNode(
                    BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(4)) == uint.MaxValue
                        ? "NilBrc" : "Brc", "BorderValue", node.StreamName, node.Offset, 8));
                break;
            case "Brc80MayBeNil" when bytes.Length == 4 &&
                BinaryPrimitives.ReadUInt32LittleEndian(bytes) != uint.MaxValue:
                node.Children.Add(new DocStructureNode("Brc80", "BorderValue",
                    node.StreamName, node.Offset, 4));
                break;
            case "Brc" when bytes.Length == 8:
                node.Children.Add(new DocStructureNode("COLORREF", "Color", node.StreamName,
                    node.Offset, 4));
                node.Children.Add(new DocStructureNode("BrcType", "BorderType", node.StreamName,
                    node.Offset + 5, 1));
                break;
            case "Brc80" when bytes.Length == 4:
                node.Children.Add(new DocStructureNode("BrcType", "BorderType", node.StreamName,
                    node.Offset + 1, 1));
                node.Children.Add(new DocStructureNode("Ico", "IndexedColor", node.StreamName,
                    node.Offset + 2, 1));
                break;
            case "Shd" when bytes.Length == 10:
                node.Children.Add(new DocStructureNode("COLORREF", "Foreground", node.StreamName,
                    node.Offset, 4));
                node.Children.Add(new DocStructureNode("COLORREF", "Background", node.StreamName,
                    node.Offset + 4, 4));
                node.Children.Add(new DocStructureNode("Ipat", "Pattern", node.StreamName,
                    node.Offset + 8, 2));
                break;
            case "Shd80" when bytes.Length == 2:
                var flags = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
                node.Attributes["foregroundIndex"] = (flags & 31).ToString(CultureInfo.InvariantCulture);
                node.Attributes["backgroundIndex"] = ((flags >> 5) & 31).ToString(CultureInfo.InvariantCulture);
                node.Attributes["pattern"] = (flags >> 10).ToString(CultureInfo.InvariantCulture);
                break;
        }
    }

    private static void AddBorders(DocStructureNode node, int start, int size, string kind)
    {
        for (var i = 0; i < Sides.Length; i++)
            node.Children.Add(new DocStructureNode(kind, Sides[i], node.StreamName,
                node.Offset + start + i * (long)size, size));
    }

    private static void AddShadingArray(DocStructureNode node, int size, string kind)
    {
        for (var i = 0; i < (node.Length - 1) / size; i++)
            node.Children.Add(new DocStructureNode(kind, $"Cell{i}", node.StreamName,
                node.Offset + 1 + i * (long)size, size));
    }
}
