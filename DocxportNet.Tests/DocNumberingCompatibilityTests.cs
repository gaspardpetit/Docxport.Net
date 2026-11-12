using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;
using DocxportNet.Visitors.Markdown;

namespace DocxportNet.Tests;

public sealed class DocNumberingCompatibilityTests
{
    [Theory]
    [InlineData("w14", "http://schemas.microsoft.com/office/word/2010/wordml")]
    [InlineData("alias", "http://schemas.microsoft.com/office/word/2010/wordml")]
    [InlineData("w15", "http://schemas.microsoft.com/office/word/2012/wordml")]
    [InlineData("w alias", "http://schemas.microsoft.com/office/word/2010/wordml")]
    public void DocUsesLegacyFallbackWhileTextKeepsCustomChoice(string requires, string extensionNamespace)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        var numbering = main.AddNewPart<NumberingDefinitionsPart>();
        numbering.Numbering = new Numbering($"""
            <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:w14="{extensionNamespace}" xmlns:w15="{extensionNamespace}" xmlns:alias="{extensionNamespace}">
              <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="7"/>
                <mc:AlternateContent><mc:Choice Requires="{requires}">
                  <w:numFmt w:val="custom" w:format="001, 002, 003, ..."/>
                </mc:Choice><mc:Fallback><w:numFmt w:val="upperRoman"/></mc:Fallback></mc:AlternateContent>
                <w:lvlText w:val="BEGIN%1END"/>
              </w:lvl></w:abstractNum>
              <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
            </w:numbering>
            """);
        main.Document = new Document(new Body(new Paragraph(
            new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
            new Run(new Text("Item0")))));
        // Materialize SDK nodes before checking that normalization leaves them intact.
        _ = numbering.Numbering.Descendants().ToArray();
        var original = numbering.Numbering.OuterXml;
        string Text(WordprocessingDocument source) => DxpExport.ExportToString(source,
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
        Assert.Contains("BEGIN007END", Text(document));
        using var binarySource = new MemoryStream();
        using (var copy = document.Clone(binarySource)) { }
        var binary = DxpDocExport.Export(binarySource.ToArray());
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Contains("BEGINVIIEND", Text(projected));
        Assert.Contains("Item0", Text(projected));
        Assert.Equal(original, numbering.Numbering.OuterXml);
        Assert.Contains("BEGIN007END", Text(document));
    }
}
