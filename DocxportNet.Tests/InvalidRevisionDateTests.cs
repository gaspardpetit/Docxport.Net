using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Doc;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class InvalidRevisionDateTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MalformedTrackedChangeDateDoesNotAbortExport(bool docExport)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var inserted = new InsertedRun(new Run(new Text("kept"))) { Id = "1", Author = "Editor" };
            var deleted = new DeletedRun(new Run(new DeletedText("removed"))) { Id = "2", Author = "Editor" };
            SetInvalidDate(inserted);
            SetInvalidDate(deleted);
            var paragraph = new Paragraph(inserted, deleted);
            var markRevision = new Deleted { Id = "3", Author = "Editor" };
            SetInvalidDate(markRevision);
            var markedParagraph = new Paragraph(new Run(new Text("mark")));
            markedParagraph.ParagraphProperties = new ParagraphProperties(
                new ParagraphMarkRunProperties(markRevision));
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(paragraph, markedParagraph));
            main.Document.Save();
        }

        var source = stream.ToArray();
        if (docExport)
        {
            var result = DxpExport.ExportToBytes(source, new DxpDocVisitor());
            Assert.Contains("kept", DxpExport.ExportToString(result,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())));
        }
        else
        {
            Assert.Contains("kept", DxpExport.ExportToString(source,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())));
        }
    }

    private static void SetInvalidDate(OpenXmlElement element) => element.SetAttribute(
        new OpenXmlAttribute("w", "date",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0000-00-00T00:00:00Z"));
}
