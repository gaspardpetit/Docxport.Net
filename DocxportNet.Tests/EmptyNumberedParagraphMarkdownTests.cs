using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class EmptyNumberedParagraphMarkdownTests
{
    [Fact]
    public void EmptyNumberedMarkerStaysInItsTableCell()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-cell-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();
                var emptyNumbered = new Paragraph(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })));
                main.Document = new Document(new Body(new Table(new TableRow(
                    new TableCell(emptyNumbered),
                    new TableCell(new Paragraph(new Run(new Text("Next cell"))))))));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Matches(@"1\s*</td>", markdown);
            Assert.DoesNotContain("1 Next cell", markdown);

            var html = DxpExport.ExportToString(path,
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
            Assert.Matches(@">1</span>\s*</p>\s*</td>", html);
            Assert.DoesNotContain(">1</span> Next cell", html);

            var plainText = DxpExport.ExportToString(path,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
            Assert.Contains("1", plainText);
            Assert.Contains("Next cell", plainText);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void PlainMarkdownKeepsNumberWhenEmptyNumberedParagraphPrecedesItsText()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbered-continuation-{Guid.NewGuid():N}.docx");
        try
        {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = document.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "%1" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
                numbering.Numbering.Save();

                Paragraph Numbered(string? text, bool tabOnly = false) => new(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
                    text is null ? (tabOnly ? new Run(new TabChar()) : null) : new Run(new Text(text)));

                main.Document = new Document(new Body(
                    Numbered("First paragraph"),
                    Numbered(null),
                    new Paragraph(new Run(new Text("Second paragraph"))),
                    Numbered(null, tabOnly: true),
                    new Paragraph(new Run(new Text("Third paragraph"))),
                    Numbered("Fourth paragraph"),
                    Numbered(null),
                    Numbered("Sixth paragraph"),
                    Numbered(null),
                    new Paragraph(),
                    new Paragraph(new Run(new Text("Separated paragraph")))));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("2 Second paragraph", markdown);
            Assert.Contains("1 First paragraph", markdown);
            Assert.Contains("3 Third paragraph", markdown);
            Assert.Contains("4 Fourth paragraph", markdown);
            Assert.Contains("5\n", markdown);
            Assert.Contains("6 Sixth paragraph", markdown);
            Assert.Contains("7\n", markdown);
            Assert.DoesNotContain("7 Separated paragraph", markdown);

            var html = DxpExport.ExportToString(path,
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
            Assert.Contains(">2</span> Second paragraph", html);
            Assert.Contains(">3</span> Third paragraph", html);
            Assert.Contains(">5</span>", html);
            Assert.Contains(">7</span>", html);
            Assert.DoesNotContain(">7</span> Separated paragraph", html);

            var plainText = DxpExport.ExportToString(path,
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()));
            Assert.Contains("2 Second paragraph", plainText);
            Assert.Contains("3 Third paragraph", plainText);
            Assert.Contains("5\n", plainText);
            Assert.Contains("7\n", plainText);
            Assert.DoesNotContain("7 Separated paragraph", plainText);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
