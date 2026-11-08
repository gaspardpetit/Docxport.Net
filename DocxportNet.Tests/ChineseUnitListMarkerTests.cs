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

public sealed class ChineseUnitListMarkerTests
{
    [Theory]
    [InlineData("chineseCountingThousand")]
    [InlineData("taiwaneseCountingThousand")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Captured from Word PDFs for standalone start overrides.
    [Theory]
    [InlineData("chineseCountingThousand", 0, true, "\u3007")]
    [InlineData("chineseCountingThousand", 1, true, "\u4e00")]
    [InlineData("chineseCountingThousand", 10, true, "\u5341")]
    [InlineData("chineseCountingThousand", 11, true, "\u5341\u4e00")]
    [InlineData("chineseCountingThousand", 21, true, "\u4e8c\u5341\u4e00")]
    [InlineData("chineseCountingThousand", 100, true, "\u4e00\u767e")]
    [InlineData("chineseCountingThousand", 101, true, "\u4e00\u767e\u3007\u4e00")]
    [InlineData("chineseCountingThousand", 110, true, "\u4e00\u767e\u4e00\u5341")]
    [InlineData("chineseCountingThousand", 1000, true, "\u4e00\u5343")]
    [InlineData("chineseCountingThousand", 1001, true, "\u4e00\u5343\u3007\u4e00")]
    [InlineData("chineseCountingThousand", 1010, true, "\u4e00\u5343\u3007\u4e00\u5341")]
    [InlineData("chineseCountingThousand", 1100, true, "\u4e00\u5343\u4e00\u767e")]
    [InlineData("chineseCountingThousand", 10000, true, "\u4e00\u4e07")]
    [InlineData("chineseCountingThousand", 10001, true, "\u4e00\u4e07\u3007\u4e00")]
    [InlineData("chineseCountingThousand", 10010, true, "\u4e00\u4e07\u3007\u4e00\u5341")]
    [InlineData("chineseCountingThousand", 10100, true, "\u4e00\u4e07\u3007\u4e00\u767e")]
    [InlineData("chineseCountingThousand", 11000, true, "\u4e00\u4e07\u4e00\u5343")]
    [InlineData("chineseCountingThousand", 100000, true, "\u4e00\u5341\u4e07")]
    [InlineData("chineseCountingThousand", 100010, true, "\u4e00\u5341\u4e07\u3007\u4e00\u5341")]
    [InlineData("chineseCountingThousand", 100100, true, "\u4e00\u5341\u4e07\u3007\u4e00\u767e")]
    [InlineData("chineseCountingThousand", 101000, true, "\u4e00\u5341\u4e07\u3007\u4e00\u5343")]
    [InlineData("chineseCountingThousand", 110000, true, "\u4e00\u5341\u4e00\u4e07")]
    [InlineData("chineseCountingThousand", 999999, true, "\u4e5d\u5341\u4e5d\u4e07\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("chineseCountingThousand", 1000000, true, "")]
    [InlineData("taiwaneseCountingThousand", 0, true, "\u96f6")]
    [InlineData("taiwaneseCountingThousand", 1, true, "\u4e00")]
    [InlineData("taiwaneseCountingThousand", 10, true, "\u5341")]
    [InlineData("taiwaneseCountingThousand", 11, true, "\u5341\u4e00")]
    [InlineData("taiwaneseCountingThousand", 21, true, "\u4e8c\u5341\u4e00")]
    [InlineData("taiwaneseCountingThousand", 100, true, "\u4e00\u767e")]
    [InlineData("taiwaneseCountingThousand", 101, true, "\u4e00\u767e\u96f6\u4e00")]
    [InlineData("taiwaneseCountingThousand", 110, true, "\u4e00\u767e\u4e00\u5341")]
    [InlineData("taiwaneseCountingThousand", 1000, true, "\u4e00\u5343")]
    [InlineData("taiwaneseCountingThousand", 1001, true, "\u4e00\u5343\u96f6\u4e00")]
    [InlineData("taiwaneseCountingThousand", 1010, true, "\u4e00\u5343\u96f6\u4e00\u5341")]
    [InlineData("taiwaneseCountingThousand", 1100, true, "\u4e00\u5343\u4e00\u767e")]
    [InlineData("taiwaneseCountingThousand", 10000, true, "\u4e00\u842c")]
    [InlineData("taiwaneseCountingThousand", 10001, true, "\u4e00\u842c\u96f6\u4e00")]
    [InlineData("taiwaneseCountingThousand", 10010, true, "\u4e00\u842c\u96f6\u4e00\u5341")]
    [InlineData("taiwaneseCountingThousand", 10100, true, "\u4e00\u842c\u96f6\u4e00\u767e")]
    [InlineData("taiwaneseCountingThousand", 11000, true, "\u4e00\u842c\u4e00\u5343")]
    [InlineData("taiwaneseCountingThousand", 100000, true, "\u4e00\u5341\u842c")]
    [InlineData("taiwaneseCountingThousand", 100010, true, "\u4e00\u5341\u842c\u96f6\u4e00\u5341")]
    [InlineData("taiwaneseCountingThousand", 100100, true, "\u4e00\u5341\u842c\u96f6\u4e00\u767e")]
    [InlineData("taiwaneseCountingThousand", 101000, true, "\u4e00\u5341\u842c\u96f6\u4e00\u5343")]
    [InlineData("taiwaneseCountingThousand", 110000, true, "\u4e00\u5341\u4e00\u842c")]
    [InlineData("taiwaneseCountingThousand", 999999, true, "\u4e5d\u5341\u4e5d\u842c\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("taiwaneseCountingThousand", 1000000, true, "")]
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
