using System.Buffers.Binary;
using System.Globalization;

namespace DocxportNet.Doc;

/// <summary>Resolves picture offsets from character property modifiers into the Data stream.</summary>
internal static class DocPictureNavigator
{
    public static void Expand(DocStructure structure, DocStructureNode operand)
    {
        if (operand.Children.Count != 0 || !operand.Attributes.TryGetValue("dataOffset", out var text) ||
            !uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out var offset)) return;
        var dataEntry = structure.Root.Children.FirstOrDefault(x => x.Kind == "Stream" && x.Name == "Data");
        if (structure.Root.Attributes.ContainsKey("sourceEncryption") &&
            dataEntry?.Length is long encryptedDataLength && offset <= encryptedDataLength - 8)
        {
            var objectHeader = structure.ReadRange("Data", offset, 8);
            var headerBytes = BinaryPrimitives.ReadUInt16LittleEndian(objectHeader);
            var objectBytes = BinaryPrimitives.ReadInt32LittleEndian(objectHeader.AsSpan(4));
            if (headerBytes == 8 && objectBytes >= 8 && objectBytes <= encryptedDataLength - offset)
            {
                var obj = new DocStructureNode("FOBJH", "EncryptedObjectHeader", "Data", offset, 8);
                obj.Attributes["objectBytes"] = objectBytes.ToString(CultureInfo.InvariantCulture);
                obj.Attributes["compressed"] = (objectHeader[2] & 1) != 0 ? "true" : "false";
                operand.Children.Add(obj);
                return;
            }
        }
        if (dataEntry?.Length is not long dataLength || offset > dataLength - 68) return;
        var header = structure.ReadRange("Data", offset, 68);
        var length = BinaryPrimitives.ReadInt32LittleEndian(header);
        var headerSize = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(4, 2));
        var metafileType = BinaryPrimitives.ReadUInt16LittleEndian(header.AsSpan(6, 2));
        if (headerSize != 68 || length < 68 || length > dataLength - offset) return;
        if (header.Skip(6).All(x => x == 0))
        {
            var nil = new DocStructureNode("NilPICFAndBinData", "FieldBinaryData", "Data",
                offset, length);
            var bodyLength = length - 68;
            if (bodyLength >= 10)
            {
                var body = structure.ReadRange("Data", offset + 68, bodyLength);
                if (BinaryPrimitives.ReadUInt32LittleEndian(body) == 0xFFFFFFFF)
                {
                    var form = new DocStructureNode("FFData", "FormField", "Data", offset + 68, bodyLength);
                    form.Children.Add(new DocStructureNode("FFDataBits", "FormFieldFlags", "Data",
                        offset + 72, 2));
                    nil.Children.Add(form);
                }
                else if (bodyLength >= 17 && (body[0] & 0xE0) == 0 &&
                    body.Skip(1).Take(16).Any(x => x != 0))
                {
                    var hyperlink = new DocStructureNode("HFD", "HyperlinkField", "Data",
                        offset + 68, bodyLength);
                    hyperlink.Children.Add(new DocStructureNode("HFDBits", "HyperlinkFlags", "Data",
                        offset + 68, 1));
                    nil.Children.Add(hyperlink);
                }
            }
            operand.Children.Add(nil);
            return;
        }
        if (metafileType is not (0x64 or 0x66)) return;

        var image = new DocStructureNode("PICFAndOfficeArtData", "InlinePicture", "Data", offset, length);
        image.Attributes["metafileType"] = $"0x{metafileType:X4}";
        var picf = new DocStructureNode("PICF", "PictureHeader", "Data", offset, 68);
        picf.Attributes["totalBytes"] = length.ToString(CultureInfo.InvariantCulture);
        picf.Children.Add(new DocStructureNode("MFPF", "MetafileFormat", "Data", offset + 6, 8));
        picf.Children.Add(new DocStructureNode("PICF_Shape", "ShapeHeader", "Data", offset + 14, 14));
        picf.Children.Add(new DocStructureNode("PICMID", "PictureMeasurements", "Data", offset + 28, 38));
        image.Children.Add(picf);
        operand.Children.Add(image);
    }
}
