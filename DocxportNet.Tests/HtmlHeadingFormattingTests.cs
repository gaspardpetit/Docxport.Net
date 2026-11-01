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

    [Fact]
    public void PlainHtmlRetainsSubtitleHeadingLevelWhenHeadingTagsAreDisabled()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(new Style(new StyleName { Val = "Subtitle" }) {
                Type = StyleValues.Paragraph,
                StyleId = "Subtitle",
            });
            styles.Styles.Save();
            main.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Subtitle" }),
                new Run(new Text("Subtitle text")))));
            main.Document.Save();
        }

        var config = DxpHtmlVisitorConfig.CreatePlainConfig() with {
            DetectHeadings = false,
            EmitParagraphMetadata = true,
        };
        var html = DxpExport.ExportToString(source.ToArray(), new DxpHtmlVisitor(config));
        Assert.Contains("<p class=\"dxp-paragraph\"", html);
        Assert.Contains("data-style-id=\"Subtitle\"", html);
        Assert.Contains("data-heading-level=\"2\"", html);
        Assert.DoesNotContain("<h2", html);
    }

    [Fact]
    public void ParagraphMetadataRetainsNumberingForQuoteAndCodeStyles()
    {
        using var source = new MemoryStream();
        using (var document = WordprocessingDocument.Create(source,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var numbering = main.AddNewPart<NumberingDefinitionsPart>();
            numbering.Numbering = new Numbering(
                new AbstractNum(new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "[%1]" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            numbering.Numbering.Save();

            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new Style(new StyleName { Val = "Quote" }) { Type = StyleValues.Paragraph, StyleId = "Quote" },
                new Style(new StyleName { Val = "Code" }) { Type = StyleValues.Paragraph, StyleId = "Code" });
            styles.Styles.Save();

            Paragraph Numbered(string style, string text) => new(
                new ParagraphProperties(new ParagraphStyleId { Val = style },
                    new NumberingProperties(new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 1 })),
                new Run(new Text(text)));
            main.Document = new Document(new Body(
                Numbered("Quote", "Quoted text"), Numbered("Code", "Code text")));
            main.Document.Save();
        }

        var config = DxpHtmlVisitorConfig.CreatePlainConfig() with { EmitParagraphMetadata = true };
        var html = DxpExport.ExportToString(source.ToArray(), new DxpHtmlVisitor(config));
        Assert.Matches("<p[^>]*data-style-id=\"Quote\"[^>]*data-num-id=\"1\"[^>]*data-num-level=\"0\"[^>]*data-marker=\"\\[1\\]\"", html);
        Assert.Matches("<pre[^>]*data-style-id=\"Code\"[^>]*data-num-id=\"1\"[^>]*data-num-level=\"0\"[^>]*data-marker=\"\\[2\\]\"", html);
        Assert.DoesNotContain("<span class=\"dxp-marker\">", html);
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
