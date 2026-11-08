using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class KoreanDigital2ListMarkerTests
{
    [Theory]
    [InlineData("koreanDigital2")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Captured from Word PDFs for standalone start overrides.
    [Theory]
    [InlineData("koreanDigital2", 0, true, "零")]
    [InlineData("koreanDigital2", 1, true, "一")]
    [InlineData("koreanDigital2", 10, true, "一零")]
    [InlineData("koreanDigital2", 11, true, "一一")]
    [InlineData("koreanDigital2", 21, true, "二一")]
    [InlineData("koreanDigital2", 100, true, "一零零")]
    [InlineData("koreanDigital2", 101, true, "一零一")]
    [InlineData("koreanDigital2", 110, true, "一一零")]
    [InlineData("koreanDigital2", 1000, true, "一零零零")]
    [InlineData("koreanDigital2", 1001, true, "一零零一")]
    [InlineData("koreanDigital2", 1010, true, "一零一零")]
    [InlineData("koreanDigital2", 1100, true, "一一零零")]
    [InlineData("koreanDigital2", 10000, true, "一零零零零")]
    [InlineData("koreanDigital2", 10001, true, "一零零零一")]
    [InlineData("koreanDigital2", 10010, true, "一零零一零")]
    [InlineData("koreanDigital2", 10100, true, "一零一零零")]
    [InlineData("koreanDigital2", 11000, true, "一一零零零")]
    [InlineData("koreanDigital2", 100000, true, "一零零零零零")]
    [InlineData("koreanDigital2", 100010, true, "一零零零一零")]
    [InlineData("koreanDigital2", 100100, true, "一零零一零零")]
    [InlineData("koreanDigital2", 101000, true, "一零一零零零")]
    [InlineData("koreanDigital2", 110000, true, "一一零零零零")]
    [InlineData("koreanDigital2", 999999, true, "九九九九九九")]
    [InlineData("koreanDigital2", 1000000, true, "")]
    public void RegionalDigitsMatchWordAtBoundaries(string format, int start, bool startOverride, string? expected)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(new AbstractNum(new MultiLevelType { Val = MultiLevelValues.Multilevel },
            new Level(new StartNumberingValue { Val = startOverride ? 1 : start },
                new NumberingFormat { Val = new NumberFormatValues(format) },
                new LevelText { Val = "BEGIN%1END" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
        if (startOverride)
            numbering.Append(new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = 0 }) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        Paragraph Item(int numId) => new(new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numId })), new Run(new Text("Body")));
        var body = new Body(Item(1));
        if (startOverride) body.Append(Item(2));
        main.Document = new Document(body);
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            var matches = Regex.Matches(output, "BEGIN(.*?)END");
            Assert.Equal(expected, matches.Count == 0 ? null : matches[matches.Count - 1].Groups[1].Value);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
