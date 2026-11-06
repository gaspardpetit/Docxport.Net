using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;

namespace DocxportNet.Tests;

public sealed class LinkedNumberingTests
{
    [Fact]
    public void WordSuppressesStyleAssociatedDefinitionWithoutItsOwnLevels()
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body());
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
            new AbstractNum(new StyleLink { Val = "Linked" }, new NumberingStyleLink { Val = "Next" }) { AbstractNumberId = 2 },
            new AbstractNum(new StyleLink { Val = "Next" }, new Level(new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal }, new LevelText { Val = "CHAIN-%1" })
                { LevelIndex = 0 }) { AbstractNumberId = 3 },
            new NumberingInstance(new AbstractNumId { Val = 2 }) { NumberID = 2 },
            new NumberingInstance(new AbstractNumId { Val = 3 }) { NumberID = 3 });
        main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
            new Style(new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 3 })))
            { Type = StyleValues.Numbering, StyleId = "Next" });
        Assert.Null(new DocxportNet.Walker.DxpNumberingResolver(doc).ResolveLevel(2, 0));
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, true)]
    public void LinkedDefinitionsFollowWordAssociationAndCounterRules(bool reverse, bool inherited, bool restart)
    {
        var path = Path.Combine(Path.GetTempPath(), $"linked-numbering-{Guid.NewGuid():N}.docx");
        try
        {
            using (var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var backing = new AbstractNum(new Level(new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "[%1]" }) { LevelIndex = 0 }) { AbstractNumberId = 2 };
                if (reverse)
                    backing.PrependChild(new StyleLink { Val = "Linked" });
                var backingInstance = new NumberingInstance(new AbstractNumId { Val = 2 }) { NumberID = 2 };
                if (restart)
                    backingInstance.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 9 }) { LevelIndex = 0 });
                main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                    new AbstractNum(new NumberingStyleLink { Val = "Linked" }) { AbstractNumberId = 1 }, backing,
                    new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 }, backingInstance);
                var linked = new Style { Type = StyleValues.Numbering, StyleId = "Linked" };
                if (inherited)
                    linked.Append(new BasedOn { Val = "Base" });
                else
                    linked.Append(new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 2 })));
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(linked,
                    new Style(new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 2 })))
                    { Type = StyleValues.Numbering, StyleId = "Base" });
                main.Document = new Document(new Body(new[] { 1, 1, 2, 1 }.Select((id, i) =>
                    new Paragraph(new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 }, new NumberingId { Val = id })),
                        new Run(new Text($"Item{i + 1}"))))));
                main.Document.Save();
            }
            var text = DxpExport.ExportToString(path,
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
            Assert.Contains(restart ? "[9] Item3" : "[1] Item3", text);
            if (reverse)
            {
                Assert.Contains("[1] Item1", text);
                Assert.Contains("[2] Item2", text);
                Assert.Contains("[3] Item4", text);
            }
            else
            {
                Assert.DoesNotContain("[1] Item1", text);
                Assert.DoesNotContain("[2] Item2", text);
                Assert.DoesNotContain("[3] Item4", text);
            }
        }
        finally { File.Delete(path); }
    }
}
