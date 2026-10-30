using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;

namespace DocxportNet.Tests;

public sealed class EmptyNumberedParagraphMarkdownTests
{
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
                    Numbered(null)));
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
            Assert.EndsWith("7", markdown.TrimEnd());
        }
        finally
        {
            File.Delete(path);
        }
    }
}
