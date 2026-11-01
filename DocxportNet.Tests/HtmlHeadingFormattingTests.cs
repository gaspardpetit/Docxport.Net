using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.Visitors.Html;

namespace DocxportNet.Tests;

public sealed class HtmlHeadingFormattingTests
{
    [Fact]
    public void PlainHtmlCanKeepHeadingAsParagraphWithStyleMetadata()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Title4LevelTitre4Niveau" },
                    new OutlineLevel { Val = 3 }),
                new Run(new Text("Heading text")))));
            main.Document.Save();
        }

        var config = DxpHtmlVisitorConfig.CreatePlainConfig() with {
            DetectHeadings = false,
            EmitParagraphMetadata = true,
        };
        var html = DxpExport.ExportToString(source.ToArray(), new DxpHtmlVisitor(config));
        Assert.Contains("<p class=\"dxp-paragraph\"", html);
        Assert.Contains("data-style-id=\"Title4LevelTitre4Niveau\"", html);
        Assert.Contains("data-heading-level=\"5\"", html);
        Assert.DoesNotContain("<h5", html);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CenteredOutlineParagraphKeepsItsActualBoldWeight(bool bold)
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new OutlineLevel { Val = 0 },
                    new Justification { Val = JustificationValues.Center }),
                new Run(bold ? new RunProperties(new Bold()) : new RunProperties(),
                    new Text("Generic heading")))));
            main.Document.Save();
        }

        var html = DxpExport.ExportToString(source.ToArray(),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
        Assert.Contains(".dxp-heading {", html);
        Assert.Contains(bold ? "<h2 class=\"dxp-heading dxp-heading-2 align-center\">" :
            "<h2 class=\"dxp-heading dxp-heading-2 align-center\" style=\"font-weight:normal;\">", html);
        Assert.DoesNotContain("<strong class=\"dxp-bold\">Generic heading</strong>", html);
    }

    [Fact]
    public void RegularOutlineHeadingCanContainBoldText()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new OutlineLevel { Val = 0 }),
                new Run(new Text("Regular ")),
                new Run(new RunProperties(new Bold()), new Text("emphasis")))));
            main.Document.Save();
        }

        var html = DxpExport.ExportToString(source.ToArray(),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
        Assert.Contains("<h2 class=\"dxp-heading dxp-heading-2\" style=\"font-weight:normal;\">", html);
        Assert.Contains("<strong class=\"dxp-bold\">emphasis</strong>", html);
    }
}
