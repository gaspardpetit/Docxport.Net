using System.Buffers.Binary;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.Doc;
using OpenMcdf;

namespace DocxportNet.Tests;

public sealed class DocLinkedPictureProjectionTests
{
    [Fact]
    public void LinkedInlinePicturePreservesExternalTargetAndExtent()
    {
        var png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z7n8AAAAASUVORK5CYII=");
        const string text = "\u0001\r";
        var story = new DocPlainTextStory(text, [], [])
        {
            Paragraphs = [new DocStoryParagraphRange(0, text.Length, DocParagraphEnd.ParagraphMark)],
            Pictures = [new DocStoryInlinePicture(0)
            {
                Payload = new DocInlinePicture(png, "image/png", 914400, 457200)
            }]
        };
        using var binary = new MemoryStream();
        DocPlainTextWriter.Write(binary, new DocPlainTextDocument(story,
            [new DocPlainTextSection(text.Length, new DocPlainTextStory?[6])]));
        const string target = "https://example.test/generic.png";
        var linkBytes = System.Text.Encoding.Unicode.GetBytes(target + "\0");
        var property = new byte[6 + linkBytes.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(property, 0xC105); // pibName.
        BinaryPrimitives.WriteUInt32LittleEndian(property.AsSpan(2), (uint)linkBytes.Length);
        linkBytes.CopyTo(property.AsSpan(6));
        var fopt = Record(0x0013, 0xF00B, property);
        var container = Record(0x000F, 0xF004, fopt);
        var linkedPicf = new byte[69 + container.Length];
        using (var compound = RootStorage.Open(binary, StorageModeFlags.LeaveOpen))
        using (var data = compound.OpenStream("Data"))
        {
            data.ReadExactly(linkedPicf.AsSpan(0, 68)); // Retain PICF dimensions.
            BinaryPrimitives.WriteInt32LittleEndian(linkedPicf, linkedPicf.Length);
            BinaryPrimitives.WriteUInt16LittleEndian(linkedPicf.AsSpan(6), 0x0066);
            container.CopyTo(linkedPicf.AsSpan(69)); // Zero-length PICF name at byte 68.
            data.Position = 0;
            data.Write(linkedPicf);
            compound.Flush();
        }

        var result = DxpDocToDocx.Project(binary.ToArray());
        using var projected = WordprocessingDocument.Open(new MemoryStream(result.DocxBytes), false);
        var main = projected.MainDocumentPart!;
        var blip = Assert.Single(main.Document!.Descendants<Blip>());
        var link = Assert.IsType<string>(blip.Link?.Value);
        Assert.Equal(new Uri(target), main.ExternalRelationships.Single(x => x.Id == link).Uri);
        Assert.Empty(main.ImageParts);
        var extent = Assert.Single(main.Document.Descendants<
            DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent>());
        Assert.Equal(914400L, extent.Cx!.Value);
        Assert.Equal(457200L, extent.Cy!.Value);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
    }

    private static byte[] Record(ushort options, ushort type, byte[] body)
    {
        var bytes = new byte[8 + body.Length];
        BinaryPrimitives.WriteUInt16LittleEndian(bytes, options);
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(2), type);
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(4), (uint)body.Length);
        body.CopyTo(bytes.AsSpan(8));
        return bytes;
    }
}
