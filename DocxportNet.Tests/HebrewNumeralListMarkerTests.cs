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

public sealed class HebrewNumeralListMarkerTests
{
    [Fact]
    public void SharedFieldFormatterRetainsDecimalFallback()
    {
        Assert.Equal("15", DxpListMarkerFormatter.FormatNumber(15, NumberFormatValues.Hebrew1));
        Assert.Equal("393", DxpListMarkerFormatter.FormatNumber(393, NumberFormatValues.Hebrew1));
    }

    // Word ListString captures: preserve logical Unicode including invisible bidi and joining controls.
    [Theory]
    [InlineData("hebrew1", -2147483648, true, "")]
    [InlineData("hebrew1", -1, true, "")]
    [InlineData("hebrew1", 0, true, "")]
    [InlineData("hebrew1", 1, true, "\u05d0")]
    [InlineData("hebrew1", 2, true, "\u05d1")]
    [InlineData("hebrew1", 3, true, "\u05d2")]
    [InlineData("hebrew1", 4, true, "\u05d3")]
    [InlineData("hebrew1", 5, true, "\u05d4")]
    [InlineData("hebrew1", 6, true, "\u05d5")]
    [InlineData("hebrew1", 7, true, "\u05d6")]
    [InlineData("hebrew1", 8, true, "\u05d7")]
    [InlineData("hebrew1", 9, true, "\u05d8")]
    [InlineData("hebrew1", 10, true, "\u05d9")]
    [InlineData("hebrew1", 11, true, "\u05d9\u05d0")]
    [InlineData("hebrew1", 12, true, "\u05d9\u05d1")]
    [InlineData("hebrew1", 13, true, "\u05d9\u05d2")]
    [InlineData("hebrew1", 14, true, "\u05d9\u05d3")]
    [InlineData("hebrew1", 15, true, "\u05d8\u05d5")]
    [InlineData("hebrew1", 16, true, "\u05d8\u05d6")]
    [InlineData("hebrew1", 17, true, "\u05d9\u05d6")]
    [InlineData("hebrew1", 18, true, "\u05d9\u05d7")]
    [InlineData("hebrew1", 19, true, "\u05d9\u05d8")]
    [InlineData("hebrew1", 20, true, "\u05db")]
    [InlineData("hebrew1", 21, true, "\u05db\u05d0")]
    [InlineData("hebrew1", 22, true, "\u05db\u05d1")]
    [InlineData("hebrew1", 23, true, "\u05db\u05d2")]
    [InlineData("hebrew1", 24, true, "\u05db\u05d3")]
    [InlineData("hebrew1", 25, true, "\u05db\u05d4")]
    [InlineData("hebrew1", 26, true, "\u05db\u05d5")]
    [InlineData("hebrew1", 27, true, "\u05db\u05d6")]
    [InlineData("hebrew1", 28, true, "\u05db\u05d7")]
    [InlineData("hebrew1", 29, true, "\u05db\u05d8")]
    [InlineData("hebrew1", 30, true, "\u05dc")]
    [InlineData("hebrew1", 31, true, "\u05dc\u05d0")]
    [InlineData("hebrew1", 32, true, "\u05dc\u05d1")]
    [InlineData("hebrew1", 33, true, "\u05dc\u05d2")]
    [InlineData("hebrew1", 34, true, "\u05dc\u05d3")]
    [InlineData("hebrew1", 35, true, "\u05dc\u05d4")]
    [InlineData("hebrew1", 36, true, "\u05dc\u05d5")]
    [InlineData("hebrew1", 37, true, "\u05dc\u05d6")]
    [InlineData("hebrew1", 38, true, "\u05dc\u05d7")]
    [InlineData("hebrew1", 39, true, "\u05dc\u05d8")]
    [InlineData("hebrew1", 40, true, "\u05de")]
    [InlineData("hebrew1", 41, true, "\u05de\u05d0")]
    [InlineData("hebrew1", 42, true, "\u05de\u05d1")]
    [InlineData("hebrew1", 43, true, "\u05de\u05d2")]
    [InlineData("hebrew1", 44, true, "\u05de\u05d3")]
    [InlineData("hebrew1", 45, true, "\u05de\u05d4")]
    [InlineData("hebrew1", 46, true, "\u05de\u05d5")]
    [InlineData("hebrew1", 47, true, "\u05de\u05d6")]
    [InlineData("hebrew1", 48, true, "\u05de\u05d7")]
    [InlineData("hebrew1", 49, true, "\u05de\u05d8")]
    [InlineData("hebrew1", 50, true, "\u05e0")]
    [InlineData("hebrew1", 51, true, "\u05e0\u05d0")]
    [InlineData("hebrew1", 52, true, "\u05e0\u05d1")]
    [InlineData("hebrew1", 53, true, "\u05e0\u05d2")]
    [InlineData("hebrew1", 54, true, "\u05e0\u05d3")]
    [InlineData("hebrew1", 55, true, "\u05e0\u05d4")]
    [InlineData("hebrew1", 56, true, "\u05e0\u05d5")]
    [InlineData("hebrew1", 57, true, "\u05e0\u05d6")]
    [InlineData("hebrew1", 58, true, "\u05e0\u05d7")]
    [InlineData("hebrew1", 59, true, "\u05e0\u05d8")]
    [InlineData("hebrew1", 60, true, "\u05e1")]
    [InlineData("hebrew1", 61, true, "\u05e1\u05d0")]
    [InlineData("hebrew1", 62, true, "\u05e1\u05d1")]
    [InlineData("hebrew1", 63, true, "\u05e1\u05d2")]
    [InlineData("hebrew1", 64, true, "\u05e1\u05d3")]
    [InlineData("hebrew1", 65, true, "\u05e1\u05d4")]
    [InlineData("hebrew1", 66, true, "\u05e1\u05d5")]
    [InlineData("hebrew1", 67, true, "\u05e1\u05d6")]
    [InlineData("hebrew1", 68, true, "\u05e1\u05d7")]
    [InlineData("hebrew1", 69, true, "\u05e1\u05d8")]
    [InlineData("hebrew1", 70, true, "\u05e2")]
    [InlineData("hebrew1", 71, true, "\u05e2\u05d0")]
    [InlineData("hebrew1", 72, true, "\u05e2\u05d1")]
    [InlineData("hebrew1", 73, true, "\u05e2\u05d2")]
    [InlineData("hebrew1", 74, true, "\u05e2\u05d3")]
    [InlineData("hebrew1", 75, true, "\u05e2\u05d4")]
    [InlineData("hebrew1", 76, true, "\u05e2\u05d5")]
    [InlineData("hebrew1", 77, true, "\u05e2\u05d6")]
    [InlineData("hebrew1", 78, true, "\u05e2\u05d7")]
    [InlineData("hebrew1", 79, true, "\u05e2\u05d8")]
    [InlineData("hebrew1", 80, true, "\u05e4")]
    [InlineData("hebrew1", 81, true, "\u05e4\u05d0")]
    [InlineData("hebrew1", 82, true, "\u05e4\u05d1")]
    [InlineData("hebrew1", 83, true, "\u05e4\u05d2")]
    [InlineData("hebrew1", 84, true, "\u05e4\u05d3")]
    [InlineData("hebrew1", 85, true, "\u05e4\u05d4")]
    [InlineData("hebrew1", 86, true, "\u05e4\u05d5")]
    [InlineData("hebrew1", 87, true, "\u05e4\u05d6")]
    [InlineData("hebrew1", 88, true, "\u05e4\u05d7")]
    [InlineData("hebrew1", 89, true, "\u05e4\u05d8")]
    [InlineData("hebrew1", 90, true, "\u05e6")]
    [InlineData("hebrew1", 91, true, "\u05e6\u05d0")]
    [InlineData("hebrew1", 92, true, "\u05e6\u05d1")]
    [InlineData("hebrew1", 93, true, "\u05e6\u05d2")]
    [InlineData("hebrew1", 94, true, "\u05e6\u05d3")]
    [InlineData("hebrew1", 95, true, "\u05e6\u05d4")]
    [InlineData("hebrew1", 96, true, "\u05e6\u05d5")]
    [InlineData("hebrew1", 97, true, "\u05e6\u05d6")]
    [InlineData("hebrew1", 98, true, "\u05e6\u05d7")]
    [InlineData("hebrew1", 99, true, "\u05e6\u05d8")]
    [InlineData("hebrew1", 100, true, "\u05e7")]
    [InlineData("hebrew1", 114, true, "\u05e7\u05d9\u05d3")]
    [InlineData("hebrew1", 115, true, "\u05e7\u05d8\u05d5")]
    [InlineData("hebrew1", 116, true, "\u05e7\u05d8\u05d6")]
    [InlineData("hebrew1", 299, true, "\u05e8\u05e6\u05d8")]
    [InlineData("hebrew1", 300, true, "\u05e9")]
    [InlineData("hebrew1", 301, true, "\u05e9\u05d0")]
    [InlineData("hebrew1", 303, true, "\u05e9\u05d2")]
    [InlineData("hebrew1", 304, true, "\u05e9\u05d3")]
    [InlineData("hebrew1", 305, true, "\u05e9\u05d4")]
    [InlineData("hebrew1", 390, true, "\u05e9\u05e6")]
    [InlineData("hebrew1", 391, true, "\u05e9\u05e6\u05d0")]
    [InlineData("hebrew1", 392, true, "\u05e9\u05e6\u05d1")]
    [InlineData("hebrew1", 393, true, "\u05d0")]
    [InlineData("hebrew1", 399, true, "\u05d6")]
    [InlineData("hebrew1", 400, true, "\u05d7")]
    [InlineData("hebrew1", 401, true, "\u05d8")]
    [InlineData("hebrew1", 539, true, "\u05e7\u05de\u05d6")]
    [InlineData("hebrew1", 540, true, "\u05e7\u05de\u05d7")]
    [InlineData("hebrew1", 541, true, "\u05e7\u05de\u05d8")]
    [InlineData("hebrew1", 557, true, "\u05e7\u05e1\u05d4")]
    [InlineData("hebrew1", 558, true, "\u05e7\u05e1\u05d5")]
    [InlineData("hebrew1", 559, true, "\u05e7\u05e1\u05d6")]
    [InlineData("hebrew1", 607, true, "\u05e8\u05d8\u05d5")]
    [InlineData("hebrew1", 608, true, "\u05e8\u05d8\u05d6")]
    [InlineData("hebrew1", 609, true, "\u05e8\u05d9\u05d6")]
    [InlineData("hebrew1", 659, true, "\u05e8\u05e1\u05d6")]
    [InlineData("hebrew1", 660, true, "\u05e8\u05e1\u05d7")]
    [InlineData("hebrew1", 661, true, "\u05e8\u05e1\u05d8")]
    [InlineData("hebrew1", 681, true, "\u05e8\u05e4\u05d8")]
    [InlineData("hebrew1", 682, true, "\u05e8\u05e6")]
    [InlineData("hebrew1", 683, true, "\u05e8\u05e6\u05d0")]
    [InlineData("hebrew1", 783, true, "\u05e9\u05e6\u05d0")]
    [InlineData("hebrew1", 784, true, "\u05e9\u05e6\u05d1")]
    [InlineData("hebrew1", 785, true, "\u05d0")]
    [InlineData("hebrew1", 839, true, "\u05e0\u05d4")]
    [InlineData("hebrew1", 840, true, "\u05e0\u05d5")]
    [InlineData("hebrew1", 841, true, "\u05e0\u05d6")]
    [InlineData("hebrew1", 867, true, "\u05e4\u05d2")]
    [InlineData("hebrew1", 868, true, "\u05e4\u05d3")]
    [InlineData("hebrew1", 869, true, "\u05e4\u05d4")]
    [InlineData("hebrew1", 999, true, "\u05e8\u05d8\u05d5")]
    [InlineData("hebrew1", 1000, true, "\u05e8\u05d8\u05d6")]
    [InlineData("hebrew1", 1001, true, "\u05e8\u05d9\u05d6")]
    [InlineData("hebrew1", 1049, true, "\u05e8\u05e1\u05d4")]
    [InlineData("hebrew1", 1050, true, "\u05e8\u05e1\u05d5")]
    [InlineData("hebrew1", 1051, true, "\u05e8\u05e1\u05d6")]
    [InlineData("hebrew1", 1084, true, "\u05e9")]
    [InlineData("hebrew1", 1085, true, "\u05e9\u05d0")]
    [InlineData("hebrew1", 1086, true, "\u05e9\u05d1")]
    [InlineData("hebrew1", 1229, true, "\u05e0\u05d2")]
    [InlineData("hebrew1", 1230, true, "\u05e0\u05d3")]
    [InlineData("hebrew1", 1231, true, "\u05e0\u05d4")]
    [InlineData("hebrew1", 1270, true, "\u05e6\u05d3")]
    [InlineData("hebrew1", 1271, true, "\u05e6\u05d4")]
    [InlineData("hebrew1", 1272, true, "\u05e6\u05d5")]
    [InlineData("hebrew1", 2459, true, "\u05e7\u05d6")]
    [InlineData("hebrew1", 2460, true, "\u05e7\u05d7")]
    [InlineData("hebrew1", 2461, true, "\u05e7\u05d8")]
    [InlineData("hebrew1", 32767, true, "\u05e8\u05dc\u05d0")]
    [InlineData("hebrew1", 32768, true, "\u05e8\u05dc\u05d1")]
    [InlineData("hebrew1", 65535, true, "\u05e2\u05d0")]
    [InlineData("hebrew1", 1000000, true, "\u05d7")]
    [InlineData("hebrew1", 2147483647, true, "\u05e8\u05dc\u05d8")]
    public void RegionalAlphabetsMatchWordAtBoundaries(string format, int start, bool startOverride, string? expected)
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
