using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Wasm;
using Xunit;

namespace DocxportNet.Tests;

public class DxpReviewTests
{
    private static byte[] CreateDocument()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.InnerXml = """
<w:p>
  <w:del w:id="35" w:author="Test"><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:delText>Old</w:delText></w:r></w:del>
  <w:ins w:id="36" w:author="Test"><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>New</w:t></w:r></w:ins>
  <w:ins w:id="37" w:author="Test"><w:r><w:t>!</w:t></w:r></w:ins>
</w:p>
""";
            main.Document = new Document(body);
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static (string Text, int Revisions) Read(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        var body = doc.MainDocumentPart!.Document.Body!;
        return (body.InnerText, body.Descendants<InsertedRun>().Count() + body.Descendants<DeletedRun>().Count());
    }

    [Fact]
    public void Html_ExposesDistinctPartQualifiedRevisionIds()
    {
        var config = DxpHtmlVisitorConfig.CreateRichConfig();
        config.TrackedChangeMode = DxpTrackedChangeMode.InlineChanges;
        string html = DxpExport.ExportToString(CreateDocument(), new DxpHtmlVisitor(config));

        Assert.Contains("class=\"dxp-deleted\" data-dxp-review-id=\"word/document.xml#35\"", html);
        Assert.Contains("class=\"dxp-inserted\" data-dxp-review-id=\"word/document.xml#36\"", html);
        Assert.Contains("class=\"dxp-inserted\" data-dxp-review-id=\"word/document.xml#37\"", html);
    }

    [Fact]
    public void Review_NullAndEmptySetsFollowBulkRules()
    {
        byte[] source = CreateDocument();
        Assert.Equal(source, DxpReview.Apply(source));
        Assert.Equal(source, DxpReview.Apply(source, [], []));
        Assert.Equal(("New!", 0), Read(DxpReview.Apply(source, null, [])));
        Assert.Equal(("Old", 0), Read(DxpReview.Apply(source, [], null)));
        Assert.Equal(("New", 0), Read(DxpReview.Apply(source, null, ["word/document.xml#37"])));
        Assert.Equal(("Old!", 0), Read(DxpReview.Apply(source, ["word/document.xml#37"], null)));
    }

    [Fact]
    public void Review_SelectsIndividualChangesAndRejectsConflicts()
    {
        byte[] source = CreateDocument();
        byte[] selected = DxpReview.Apply(source,
            ["word/document.xml#35", "word/document.xml#36"], ["word/document.xml#37"]);
        Assert.Equal(("New", 0), Read(selected));
        Assert.Equal(("New!", 1), Read(DxpReview.Apply(source,
            ["word/document.xml#35", "word/document.xml#36"], [])));
        Assert.Throws<ArgumentException>(() => DxpReview.Apply(source,
            ["word/document.xml#35"], ["word/document.xml#35"]));
        Assert.Throws<ArgumentException>(() => DxpReview.Apply(source,
            ["word/document.xml#999"], []));
    }

    [Fact]
    public void BrowserReview_UsesTheSameSelections()
    {
        var result = BrowserExports.ReviewDocxForTests(CreateDocument(), new BrowserReviewRequest {
            Accept = ["word/document.xml#35", "word/document.xml#36"],
            Reject = ["word/document.xml#37"]
        });
        Assert.Equal(("New", 0), Read(result));
    }

    [Fact]
    public void Review_HandlesPropertyAndTableRowRevisions()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.InnerXml = """
<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="40" w:author="Test"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>Styled</w:t></w:r></w:p>
<w:tbl><w:tr><w:trPr><w:del w:id="41" w:author="Test"/></w:trPr><w:tc><w:p><w:r><w:t>Row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
""";
            main.Document = new Document(body);
            main.Document.Save();
        }

        byte[] source = stream.ToArray();
        using var accepted = WordprocessingDocument.Open(new MemoryStream(DxpReview.Apply(source, null, [])), false);
        Assert.NotNull(accepted.MainDocumentPart!.Document.Descendants<Bold>().SingleOrDefault());
        Assert.Empty(accepted.MainDocumentPart.Document.Descendants<RunPropertiesChange>());
        Assert.Empty(accepted.MainDocumentPart.Document.Descendants<TableRow>());

        using var rejected = WordprocessingDocument.Open(new MemoryStream(DxpReview.Apply(source, [], null)), false);
        Assert.NotNull(rejected.MainDocumentPart!.Document.Descendants<Italic>().SingleOrDefault());
        Assert.Empty(rejected.MainDocumentPart.Document.Descendants<Bold>());
        Assert.Single(rejected.MainDocumentPart.Document.Descendants<TableRow>());
    }
}
