using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Walker;

namespace DocxportNet.Tests;

public sealed class ParagraphNumberingInstanceTests
{
    [Fact]
    public void StartOverrideStillAppliesWhenInstanceBeginsAtNumberedLevel()
    {
        var tracker = new DxpListTracker();

        Assert.Equal(1, tracker.NextIndex(1, 7, 0, 1, false));
        Assert.Equal(5, tracker.NextIndex(8, 7, 0, 5, true));
    }

    [Fact]
    public void PlainMarkdownContinuesParagraphNumberAfterHeadingUsesSameInstance()
    {
        var path = Path.Combine(Path.GetTempPath(), $"numbering-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var numbering = main.AddNewPart<NumberingDefinitionsPart>();
                numbering.Numbering = new Numbering(
                    new AbstractNum(
                        new Level(
                            new StartNumberingValue { Val = 88 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "[%1]" }) { LevelIndex = 0 },
                        new Level(
                            new StartNumberingValue { Val = 1 },
                            new NumberingFormat { Val = NumberFormatValues.Decimal },
                            new LevelText { Val = "(%4)" }) { LevelIndex = 3 })
                    { AbstractNumberId = 7 },
                    new NumberingInstance(new AbstractNumId { Val = 7 }) { NumberID = 1 },
                    new NumberingInstance(
                        new AbstractNumId { Val = 7 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 },
                        new LevelOverride(new StartOverrideNumberingValue { Val = 2 }) { LevelIndex = 3 })
                    { NumberID = 8 });
                numbering.Numbering.Save();

                var styles = main.AddNewPart<StyleDefinitionsPart>();
                styles.Styles = new Styles(new Style(
                    new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 1 })))
                { Type = StyleValues.Paragraph, StyleId = "NumberedBody" });
                styles.Styles.Save();

                main.Document = new Document(new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new NumberingProperties(new NumberingLevelReference { Val = 3 },
                                new NumberingId { Val = 8 })),
                        new Run(new Text("Analysis"))),
                    new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "NumberedBody" }),
                        new Run(new Text("First paragraph"))),
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId { Val = "NumberedBody" },
                            new NumberingProperties(new NumberingLevelReference { Val = 0 },
                                new NumberingId { Val = 8 })),
                        new Run(new Text("Second paragraph")))));
                main.Document.Save();
            }

            var markdown = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains("[88] First paragraph", markdown);
            Assert.Contains("[89] Second paragraph", markdown);
            Assert.DoesNotContain("[1] Second paragraph", markdown);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
