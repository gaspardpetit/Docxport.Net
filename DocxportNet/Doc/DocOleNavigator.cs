using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Locates the small ObjInfo record in an ObjectPool child storage.</summary>
internal static class DocOleNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode stream)
    {
        if (stream.Children.Count != 0 || stream.Name != "ObjInfo" ||
            stream.StreamName == null || stream.Length is not >= 4 ||
            !stream.StreamName.StartsWith("ObjectPool/", StringComparison.Ordinal)) return;
        var size = Math.Min(stream.Length.Value, 6L);
        var bytes = structure.ReadRange(stream.StreamName, 0, checked((int)size));
        var odt = new DocStructureNode("ODT", "ObjectDescriptor", stream.StreamName, 0, size);
        var flags = new DocStructureNode("ODTPersist1", "ObjectFlags", stream.StreamName, 0, 2);
        flags.Attributes["value"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes)
            .ToString(CultureInfo.InvariantCulture);
        odt.Children.Add(flags);
        odt.Attributes["clipboardFormat"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(2))
            .ToString(CultureInfo.InvariantCulture);
        if (size == 6)
        {
            var more = new DocStructureNode("ODTPersist2", "PresentationFlags",
                stream.StreamName, 4, 2);
            more.Attributes["value"] = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(4))
                .ToString(CultureInfo.InvariantCulture);
            odt.Children.Add(more);
        }
        stream.Children.Add(odt);
    }
}
