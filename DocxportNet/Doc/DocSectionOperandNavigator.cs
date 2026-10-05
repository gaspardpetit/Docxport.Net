using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes nested section-property revision and page-border fields.</summary>
internal static class DocSectionOperandNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length == null) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value,
            checked((int)node.Length.Value));
        if (node.Kind == "PropRMarkOperand" && bytes.Length == 8 && bytes[0] == 7)
        {
            node.Children.Add(new DocStructureNode("PropRMark", "RevisionMark",
                node.StreamName, node.Offset + 1, 7));
        }
        else if (node.Kind == "PropRMark" && bytes.Length == 7)
        {
            node.Attributes["present"] = bytes[0] == 1 ? "true" : "false";
            node.Attributes["authorIndex"] = BinaryPrimitives.ReadInt16LittleEndian(
                bytes.AsSpan(1, 2)).ToString(CultureInfo.InvariantCulture);
            node.Children.Add(new DocStructureNode("DTTM", "RevisionTime",
                node.StreamName, node.Offset + 3, 4));
        }
        else if (node.Kind == "SPgbPropOperand" && bytes.Length == 2)
        {
            var fields = bytes[0];
            Add("PgbApplyTo", "Pages", fields & 7);
            Add("PgbPageDepth", "Depth", (fields >> 3) & 3);
            Add("PgbOffsetFrom", "OffsetBasis", (fields >> 5) & 7);
        }

        void Add(string kind, string name, int value)
        {
            var child = new DocStructureNode(kind, name, node.StreamName, node.Offset, 1);
            child.Attributes["value"] = value.ToString(CultureInfo.InvariantCulture);
            node.Children.Add(child);
        }
    }
}
