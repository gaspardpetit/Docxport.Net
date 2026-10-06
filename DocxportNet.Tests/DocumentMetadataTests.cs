using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Wasm;
using W = DocumentFormat.OpenXml.Wordprocessing;
using E = DocumentFormat.OpenXml.ExtendedProperties;

namespace DocxportNet.Tests;

public sealed class DocumentMetadataTests
{
    [Fact]
    public void DocxMetadataIncludesPropertiesLanguageRevisionsAndComments()
    {
        var bytes = CreateDocx();
        var metadata = DxpMetadata.Inspect(bytes);
        Assert.Equal("Generic title", metadata.CoreProperties.Title);
        Assert.Equal("Writer", metadata.CoreProperties.Creator);
        Assert.Equal("2020-01-02T03:04:05.0000000Z", metadata.CoreProperties.Created);
        Assert.Equal("2", metadata.ExtendedProperties?.Pages);
        Assert.Equal("Generic app", metadata.ExtendedProperties?.Application);
        Assert.True(metadata.HasTrackedChanges);
        Assert.True(metadata.HasComments);
        Assert.Equal(new[] { "fr", "en" }, metadata.Language!.Select(item => item.Code));
        Assert.Equal(new[] { .7, .3 }, metadata.Language.Select(item => item.Ratio));

        var browser = BrowserExports.InspectForTests(bytes);
        Assert.Equal(metadata.CoreProperties, browser.CoreProperties);
        Assert.Equal(metadata.ExtendedProperties, browser.ExtendedProperties);
        Assert.True(browser.HasTrackedChanges);
        Assert.True(browser.HasComments);
        using var json = JsonDocument.Parse(BrowserExports.Inspect(bytes));
        Assert.Equal("Generic title", json.RootElement.GetProperty("coreProperties")
            .GetProperty("title").GetString());
        Assert.Equal("fr", json.RootElement.GetProperty("language")[0]
            .GetProperty("code").GetString());
        Assert.True(json.RootElement.GetProperty("hasComments").GetBoolean());
    }

    [Fact]
    public void BinaryDocExposesMetadataWithoutProjectingToDocx()
    {
        var binary = DxpDocExport.Export(CreateDocx());
        var metadata = DxpMetadata.Inspect(binary);
        Assert.Equal("Generic title", metadata.CoreProperties.Title);
        Assert.Equal("Writer", metadata.CoreProperties.Creator);
        Assert.Equal("2", metadata.ExtendedProperties?.Pages);
        Assert.True(metadata.HasTrackedChanges);
        Assert.Equal("fr", metadata.Language![0].Code);
        Assert.Equal(metadata.HasComments, BrowserExports.InspectForTests(binary).HasComments);
        Assert.Equal(metadata.CoreProperties, BrowserExports.InspectForTests(binary).CoreProperties);
    }

    [Fact]
    public void DocDescriptionPropertyDoesNotImplyAComment()
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "DocxportNet.Tests", "Fixtures", "Doc",
            "WordDocPropertyCommentsBodyFooter.doc"));
        var metadata = DxpMetadata.Inspect(path);
        Assert.False(metadata.HasComments);
        Assert.NotNull(metadata.CoreProperties.Description);
        Assert.False(BrowserExports.InspectForTests(File.ReadAllBytes(path)).HasComments);
    }

    private static byte[] CreateDocx()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            document.PackageProperties.Title = "Generic title";
            document.PackageProperties.Creator = "Writer";
            document.PackageProperties.Created = new DateTime(2020, 1, 2, 3, 4, 5, DateTimeKind.Utc);
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new Run(new RunProperties(new Languages { Val = "fr-CA" }), new Text("Bonjour")),
                new W.InsertedRun(new Run(new RunProperties(new Languages { Val = "en-US" }),
                    new Text("yes"))) { Id = "1", Author = "Writer" },
                new Run(new CommentReference { Id = "0" }))));
            main.Document.Save();
            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.Comments = new Comments(new Comment(new Paragraph(new Run(new Text("Note"))))
                { Id = "0", Author = "Writer" });
            comments.Comments.Save();
            var extended = document.AddExtendedFilePropertiesPart();
            extended.Properties = new E.Properties(new E.Application("Generic app"),
                new E.Pages("2"));
            extended.Properties.Save();
        }
        return stream.ToArray();
    }
}
