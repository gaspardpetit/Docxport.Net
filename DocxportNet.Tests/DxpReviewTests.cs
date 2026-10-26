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

    private static byte[] CreateReviewDocument(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            body.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            body.InnerXml = bodyXml;
            main.Document = new Document(body);
            main.Document.Save();
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("del", true)]
    [InlineData("ins", false)]
    public void Review_RemovesSelectedParagraphBreak(string mark, bool accept)
    {
        byte[] source = CreateReviewDocument($"""
<w:p><w:pPr><w:rPr><w:{mark} w:id="50" w:author="Test"/></w:rPr></w:pPr><w:r><w:t>First</w:t></w:r></w:p>
<w:p><w:r><w:t>Second</w:t></w:r></w:p>
""");
        byte[] result = accept
            ? DxpReview.Apply(source, ["word/document.xml#50"], [])
            : DxpReview.Apply(source, [], ["word/document.xml#50"]);
        using var doc = WordprocessingDocument.Open(new MemoryStream(result), false);
        var paragraphs = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToArray();
        Assert.Single(paragraphs);
        Assert.Equal("FirstSecond", paragraphs[0].InnerText);
    }

    [Fact]
    public void Review_HandlesTableCellInsertionsDeletionsAndMerges()
    {
        byte[] source = CreateReviewDocument("""
<w:tbl><w:tr>
<w:tc><w:tcPr><w:cellIns w:id="60" w:author="Test"/></w:tcPr><w:p><w:r><w:t>Inserted</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:cellDel w:id="61" w:author="Test"/></w:tcPr><w:p><w:r><w:t>Deleted</w:t></w:r></w:p></w:tc>
<w:tc><w:tcPr><w:vMerge w:val="continue"/><w:cellMerge w:id="62" w:author="Test" w:vMerge="cont" w:vMergeOrig="rest"/></w:tcPr><w:p/></w:tc>
</w:tr></w:tbl>
""");
        using var accepted = WordprocessingDocument.Open(new MemoryStream(DxpReview.Apply(source, null, [])), false);
        var acceptedCells = accepted.MainDocumentPart!.Document.Descendants<TableCell>().ToArray();
        Assert.Equal(2, acceptedCells.Length);
        Assert.Equal("Inserted", acceptedCells[0].InnerText);
        Assert.NotNull(acceptedCells[1].TableCellProperties!.GetFirstChild<VerticalMerge>());
        Assert.Empty(accepted.MainDocumentPart.Document.Descendants<CellMerge>());

        using var rejected = WordprocessingDocument.Open(new MemoryStream(DxpReview.Apply(source, [], null)), false);
        var rejectedCells = rejected.MainDocumentPart!.Document.Descendants<TableCell>().ToArray();
        Assert.Equal(2, rejectedCells.Length);
        Assert.Equal("Deleted", rejectedCells[0].InnerText);
        Assert.Null(rejectedCells[1].TableCellProperties!.GetFirstChild<VerticalMerge>());
        Assert.Empty(rejected.MainDocumentPart.Document.Descendants<CellMerge>());

        using var selected = WordprocessingDocument.Open(new MemoryStream(DxpReview.Apply(source,
            ["word/document.xml#60", "word/document.xml#62"], ["word/document.xml#61"])), false);
        Assert.Equal(3, selected.MainDocumentPart!.Document.Descendants<TableCell>().Count());
        Assert.Empty(selected.MainDocumentPart.Document.Descendants<CellMerge>());
    }

    [Fact]
    public void Review_RestoresDeletedFieldInstructions()
    {
        byte[] source = CreateReviewDocument("""
<w:p><w:del w:id="70" w:author="Test"><w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r></w:del></w:p>
""");
        using var doc = WordprocessingDocument.Open(new MemoryStream(DxpReview.Apply(source,
            [], ["word/document.xml#70"])), false);
        var field = Assert.Single(doc.MainDocumentPart!.Document.Descendants<FieldCode>());
        Assert.Equal(" PAGE ", field.Text);
        Assert.Equal(SpaceProcessingModeValues.Preserve, field.Space!.Value);
        Assert.Empty(doc.MainDocumentPart.Document.Descendants<DeletedFieldCode>());
    }
}
