using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Doc;

namespace DocxportNet.Tests;

public sealed class DocProjectionDefaultsTests
{
    [Fact]
    public void RunDefaultsWriteParagraphDefaultWithoutNormalParagraphFormatting()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream,
            WordprocessingDocumentType.Document, true);
        var main = document.AddMainDocumentPart();

        DocToDocxProjector.WriteStyles(main, Array.Empty<DocStyleDefinition>(),
            new DocCharacterFormatting(null, null, 24));

        var defaults = main.StyleDefinitionsPart!.Styles!.DocDefaults;
        Assert.NotNull(defaults?.RunPropertiesDefault);
        Assert.NotNull(defaults.ParagraphPropertiesDefault);
        Assert.Empty(defaults.ParagraphPropertiesDefault.ChildElements);
    }
}
