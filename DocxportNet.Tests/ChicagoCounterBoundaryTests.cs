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

public sealed class ChicagoCounterBoundaryTests
{
    [Fact]
    public void SharedFieldFormatterRetainsItsRepetitionLimit()
    {
        Assert.Equal(new string('*', 31), DxpListMarkerFormatter.FormatNumber(121, NumberFormatValues.Chicago));
        Assert.Equal("1025", DxpListMarkerFormatter.FormatNumber(1025, NumberFormatValues.Chicago));
    }

    // Captured from visible Word PDFs, including negative abstract starts vs start overrides.
    [Theory]
    [InlineData("chicago", 0, false, "")]
    [InlineData("chicago", 1, false, "*")]
    [InlineData("chicago", 4, false, "\u00a7")]
    [InlineData("chicago", 5, false, "**")]
    [InlineData("chicago", 119, false, "\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021")]
    [InlineData("chicago", 120, false, "\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7")]
    [InlineData("chicago", 121, false, "*")]
    [InlineData("chicago", 124, false, "\u00a7")]
    [InlineData("chicago", 1024, false, "\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7")]
    [InlineData("chicago", 1025, false, "*****************")]
    [InlineData("chicago", 2147483647, false, "\u2021\u2021")]
    [InlineData("chicago", -2147483648, false, null)]
    [InlineData("chicago", -1, false, null)]
    [InlineData("chicago", -2147483648, true, "")]
    [InlineData("chicago", -1, true, "")]
    public void RepeatedMarkersMatchWordAtBoundaries(string format, int start, bool startOverride, string? expected)
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
