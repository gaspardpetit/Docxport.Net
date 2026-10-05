using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

internal static class DocTemplateCodeNavigator
{
    public static DocStructureNode Create(string stream, long offset, ReadOnlySpan<byte> bytes, string name)
    {
        var value = BinaryPrimitives.ReadUInt32LittleEndian(bytes);
        var builtIn = (value & 1) != 0;
        var node = new DocStructureNode("Tplc", name, stream, offset, 4);
        var variant = new DocStructureNode(builtIn ? "TplcBuildIn" : "TplcUser",
            builtIn ? "BuiltInTemplate" : "UserTemplate", stream, offset, 4);
        if (builtIn)
        {
            variant.Attributes["formatIndex"] = ((value >> 1) & 0x7FFF)
                .ToString(CultureInfo.InvariantCulture);
            variant.Attributes["languageId"] = (value >> 16).ToString(CultureInfo.InvariantCulture);
        }
        else
            variant.Attributes["randomId"] = (value >> 1).ToString(CultureInfo.InvariantCulture);
        node.Children.Add(variant);
        return node;
    }
}
