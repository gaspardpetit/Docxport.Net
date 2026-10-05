using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

internal static class DocLegacyNavigator
{
    public static void ExpandSelection(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length is not long length || length < 26) return;
        var bytes = structure.ReadRange(location.StreamName, location.Offset.Value, checked((int)length));
        var flags = BinaryPrimitives.ReadUInt16LittleEndian(bytes);
        var selection = new DocStructureNode("Selsf", "LastSelection", location.StreamName,
            location.Offset, length);
        var variant = (flags & (1 << 11)) != 0 ? "TableSel" :
            (flags & (1 << 13)) != 0 ? "BlockSel" : null;
        if (variant != null)
            selection.Children.Add(new DocStructureNode(variant, "SelectedCellsOrBlock",
                location.StreamName, location.Offset + 16, 4));
        var style = new DocStructureNode("Sty", "SelectionType", location.StreamName,
            location.Offset + 24, 2);
        style.Attributes["value"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(24))
            .ToString(CultureInfo.InvariantCulture);
        selection.Children.Add(style);
        location.Children.Add(selection);
    }

    public static void ExpandAuthorFilter(DocStructure structure, DocStructureNode location)
    {
        if (location.Children.Count != 0 || location.StreamName == null ||
            location.Offset == null || location.Length < 4) return;
        var count = BinaryPrimitives.ReadInt32LittleEndian(
            structure.ReadRange(location.StreamName, location.Offset.Value, 4));
        if (count < 0 || 4L + count * 2L != location.Length) return;
        var afd = new DocStructureNode("Afd", "HiddenAuthors", location.StreamName,
            location.Offset, location.Length);
        afd.Attributes["authorCount"] = count.ToString(CultureInfo.InvariantCulture);
        location.Children.Add(afd);
    }
}
