using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;

namespace DocxportNet.Tests;

public sealed class StyleNumberingInheritanceTests
{
    [Theory]
    [InlineData(1, null, 1, 1, null, null, false)]
    [InlineData(null, 1, 1, null, null, null, false)]
    [InlineData(null, null, null, 1, 1, null, false)]
    [InlineData(null, null, 1, null, null, 1, false)]
    [InlineData(null, null, 1, 1, null, null, false)]
    [InlineData(null, null, 1, 1, null, null, true)]
    public void WordCombinesNumberingPropertiesAcrossStyleSources(int? directId, int? directLevel,
        int? styleId, int? styleLevel, int? baseId, int? baseLevel, bool defaultStyle)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        NumberingProperties Num(int? id, int? level) => new(
            level.HasValue ? new NumberingLevelReference { Val = level.Value } : null,
            id.HasValue ? new NumberingId { Val = id.Value } : null);
        var bodyStyle = new Style(new StyleParagraphProperties(Num(styleId, styleLevel)))
            { Type = StyleValues.Paragraph, StyleId = "Body", Default = defaultStyle };
        if (baseId.HasValue || baseLevel.HasValue)
            bodyStyle.PrependChild(new BasedOn { Val = "Base" });
        main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(bodyStyle,
            new Style(new StyleParagraphProperties(Num(baseId, baseLevel))) { Type = StyleValues.Paragraph, StyleId = "Base" });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
            new AbstractNum(Enumerable.Range(0, 2).Select(level => new Level(
                new StartNumberingValue { Val = 1 }, new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = $"L{level}-%{level + 1}" }) { LevelIndex = level })) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
        var properties = new ParagraphProperties();
        if (!defaultStyle)
            properties.Append(new ParagraphStyleId { Val = "Body" });
        if (directId.HasValue || directLevel.HasValue)
            properties.Append(Num(directId, directLevel));
        main.Document = new Document(new Body(new Paragraph(properties, new Run(new Text("Styled")))));
        var output = DxpExport.ExportToString(doc, new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()));
        Assert.Contains("L1-1 Styled", output);
    }
}
