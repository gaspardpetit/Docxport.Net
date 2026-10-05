using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Indexes tab-stop deletions, additions, and tab descriptor fields.</summary>
internal static class DocTabOperandNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode node)
    {
        if (node.Children.Count != 0 || node.StreamName == null ||
            node.Offset == null || node.Length == null) return;
        var bytes = structure.ReadRange(node.StreamName, node.Offset.Value,
            checked((int)node.Length.Value));
        if (node.Kind == "PChgTabsPapxOperand")
        {
            if (bytes.Length < 3 || bytes[0] + 1 != bytes.Length) return;
            var deleted = bytes[1];
            if (deleted > 64) return;
            var deleteLength = 1 + 2 * deleted;
            if (1 + deleteLength >= bytes.Length) return;
            var added = bytes[1 + deleteLength];
            if (added > 64 || 2 + deleteLength + 3 * added != bytes.Length) return;
            node.Children.Add(new DocStructureNode("PChgTabsDel", "RemovedTabStops",
                node.StreamName, node.Offset + 1, deleteLength));
            node.Children.Add(new DocStructureNode("PChgTabsAdd", "AddedTabStops",
                node.StreamName, node.Offset + 1 + deleteLength, 1 + 3 * added));
        }
        else if (node.Kind == "PChgTabsOperand")
        {
            if (bytes.Length < 3 || bytes[0] == 255 || bytes[0] + 1 != bytes.Length) return;
            var deleted = bytes[1];
            if (deleted > 64) return;
            var deleteLength = 1 + 4 * deleted;
            if (1 + deleteLength >= bytes.Length) return;
            var added = bytes[1 + deleteLength];
            if (added > 64 || 2 + deleteLength + 3 * added != bytes.Length) return;
            node.Children.Add(new DocStructureNode("PChgTabsDelClose", "RemovedTabStops",
                node.StreamName, node.Offset + 1, deleteLength));
            node.Children.Add(new DocStructureNode("PChgTabsAdd", "AddedTabStops",
                node.StreamName, node.Offset + 1 + deleteLength, 1 + 3 * added));
        }
        else if (node.Kind == "PChgTabsDel" && bytes.Length >= 1)
        {
            var count = bytes[0];
            if (bytes.Length != 1 + 2 * count) return;
            node.Attributes["tabCount"] = count.ToString(CultureInfo.InvariantCulture);
            for (var i = 0; i < count; i++)
                node.Children.Add(new DocStructureNode("XAS", $"Position{i}",
                    node.StreamName, node.Offset + 1 + i * 2L, 2));
        }
        else if (node.Kind == "PChgTabsDelClose" && bytes.Length >= 1)
        {
            var count = bytes[0];
            if (bytes.Length != 1 + 4 * count) return;
            node.Attributes["tabCount"] = count.ToString(CultureInfo.InvariantCulture);
            for (var i = 0; i < count; i++)
                node.Children.Add(new DocStructureNode("XAS_plusOne", $"CloseRange{i}",
                    node.StreamName, node.Offset + 1 + 2 * count + i * 2L, 2));
        }
        else if (node.Kind == "PChgTabsAdd" && bytes.Length >= 1)
        {
            var count = bytes[0];
            if (bytes.Length != 1 + 3 * count) return;
            node.Attributes["tabCount"] = count.ToString(CultureInfo.InvariantCulture);
            for (var i = 0; i < count; i++)
            {
                node.Children.Add(new DocStructureNode("XAS", $"Position{i}",
                    node.StreamName, node.Offset + 1 + i * 2L, 2));
                node.Children.Add(new DocStructureNode("TBD", $"Descriptor{i}",
                    node.StreamName, node.Offset + 1 + 2 * count + i, 1));
            }
        }
        else if (node.Kind == "TBD" && bytes.Length == 1)
        {
            var justification = new DocStructureNode("TabJC", "Alignment", node.StreamName,
                node.Offset, 1);
            justification.Attributes["value"] = (bytes[0] & 7).ToString(CultureInfo.InvariantCulture);
            node.Children.Add(justification);
            var leader = new DocStructureNode("TabLC", "Leader", node.StreamName,
                node.Offset, 1);
            leader.Attributes["value"] = ((bytes[0] >> 3) & 7).ToString(CultureInfo.InvariantCulture);
            node.Children.Add(leader);
        }
    }
}
