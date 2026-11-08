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

public sealed class HebrewAlphabetListMarkerTests
{
    [Fact]
    public void SharedFieldFormatterRetainsDecimalFallback()
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, NumberFormatValues.Hebrew2));
        Assert.Equal("393", DxpListMarkerFormatter.FormatNumber(393, NumberFormatValues.Hebrew2));
    }

    // Word ListString captures: preserve logical Unicode including invisible bidi and joining controls.
    [Theory]
    [InlineData("hebrew2", -2147483648, true, "")]
    [InlineData("hebrew2", -1, true, "")]
    [InlineData("hebrew2", 0, true, "")]
    [InlineData("hebrew2", 1, true, "\u200f\u05d0")]
    [InlineData("hebrew2", 2, true, "\u200f\u05d1")]
    [InlineData("hebrew2", 3, true, "\u200f\u05d2")]
    [InlineData("hebrew2", 4, true, "\u200f\u05d3")]
    [InlineData("hebrew2", 5, true, "\u200f\u05d4")]
    [InlineData("hebrew2", 6, true, "\u200f\u05d5")]
    [InlineData("hebrew2", 7, true, "\u200f\u05d6")]
    [InlineData("hebrew2", 8, true, "\u200f\u05d7")]
    [InlineData("hebrew2", 9, true, "\u200f\u05d8")]
    [InlineData("hebrew2", 10, true, "\u200f\u05d9")]
    [InlineData("hebrew2", 11, true, "\u200f\u05db")]
    [InlineData("hebrew2", 12, true, "\u200f\u05dc")]
    [InlineData("hebrew2", 13, true, "\u200f\u05de")]
    [InlineData("hebrew2", 14, true, "\u200f\u05e0")]
    [InlineData("hebrew2", 15, true, "\u200f\u05e1")]
    [InlineData("hebrew2", 16, true, "\u200f\u05e2")]
    [InlineData("hebrew2", 17, true, "\u200f\u05e4")]
    [InlineData("hebrew2", 18, true, "\u200f\u05e6")]
    [InlineData("hebrew2", 19, true, "\u200f\u05e7")]
    [InlineData("hebrew2", 20, true, "\u200f\u05e8")]
    [InlineData("hebrew2", 21, true, "\u200f\u05e9")]
    [InlineData("hebrew2", 22, true, "\u200f\u05ea")]
    [InlineData("hebrew2", 23, true, "\u200f\u05ea\u05d0")]
    [InlineData("hebrew2", 24, true, "\u200f\u05ea\u05d1")]
    [InlineData("hebrew2", 25, true, "\u200f\u05ea\u05d2")]
    [InlineData("hebrew2", 26, true, "\u200f\u05ea\u05d3")]
    [InlineData("hebrew2", 27, true, "\u200f\u05ea\u05d4")]
    [InlineData("hebrew2", 28, true, "\u200f\u05ea\u05d5")]
    [InlineData("hebrew2", 29, true, "\u200f\u05ea\u05d6")]
    [InlineData("hebrew2", 30, true, "\u200f\u05ea\u05d7")]
    [InlineData("hebrew2", 31, true, "\u200f\u05ea\u05d8")]
    [InlineData("hebrew2", 32, true, "\u200f\u05ea\u05d9")]
    [InlineData("hebrew2", 33, true, "\u200f\u05ea\u05db")]
    [InlineData("hebrew2", 34, true, "\u200f\u05ea\u05dc")]
    [InlineData("hebrew2", 35, true, "\u200f\u05ea\u05de")]
    [InlineData("hebrew2", 36, true, "\u200f\u05ea\u05e0")]
    [InlineData("hebrew2", 37, true, "\u200f\u05ea\u05e1")]
    [InlineData("hebrew2", 38, true, "\u200f\u05ea\u05e2")]
    [InlineData("hebrew2", 39, true, "\u200f\u05ea\u05e4")]
    [InlineData("hebrew2", 40, true, "\u200f\u05ea\u05e6")]
    [InlineData("hebrew2", 41, true, "\u200f\u05ea\u05e7")]
    [InlineData("hebrew2", 42, true, "\u200f\u05ea\u05e8")]
    [InlineData("hebrew2", 43, true, "\u200f\u05ea\u05e9")]
    [InlineData("hebrew2", 44, true, "\u200f\u05ea\u05ea")]
    [InlineData("hebrew2", 45, true, "\u200f\u05ea\u05ea\u05d0")]
    [InlineData("hebrew2", 46, true, "\u200f\u05ea\u05ea\u05d1")]
    [InlineData("hebrew2", 47, true, "\u200f\u05ea\u05ea\u05d2")]
    [InlineData("hebrew2", 48, true, "\u200f\u05ea\u05ea\u05d3")]
    [InlineData("hebrew2", 49, true, "\u200f\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 50, true, "\u200f\u05ea\u05ea\u05d5")]
    [InlineData("hebrew2", 51, true, "\u200f\u05ea\u05ea\u05d6")]
    [InlineData("hebrew2", 52, true, "\u200f\u05ea\u05ea\u05d7")]
    [InlineData("hebrew2", 53, true, "\u200f\u05ea\u05ea\u05d8")]
    [InlineData("hebrew2", 54, true, "\u200f\u05ea\u05ea\u05d9")]
    [InlineData("hebrew2", 55, true, "\u200f\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 56, true, "\u200f\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 57, true, "\u200f\u05ea\u05ea\u05de")]
    [InlineData("hebrew2", 58, true, "\u200f\u05ea\u05ea\u05e0")]
    [InlineData("hebrew2", 59, true, "\u200f\u05ea\u05ea\u05e1")]
    [InlineData("hebrew2", 60, true, "\u200f\u05ea\u05ea\u05e2")]
    [InlineData("hebrew2", 61, true, "\u200f\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 62, true, "\u200f\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 63, true, "\u200f\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 64, true, "\u200f\u05ea\u05ea\u05e8")]
    [InlineData("hebrew2", 65, true, "\u200f\u05ea\u05ea\u05e9")]
    [InlineData("hebrew2", 66, true, "\u200f\u05ea\u05ea\u05ea")]
    [InlineData("hebrew2", 67, true, "\u200f\u05ea\u05ea\u05ea\u05d0")]
    [InlineData("hebrew2", 68, true, "\u200f\u05ea\u05ea\u05ea\u05d1")]
    [InlineData("hebrew2", 69, true, "\u200f\u05ea\u05ea\u05ea\u05d2")]
    [InlineData("hebrew2", 70, true, "\u200f\u05ea\u05ea\u05ea\u05d3")]
    [InlineData("hebrew2", 71, true, "\u200f\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 72, true, "\u200f\u05ea\u05ea\u05ea\u05d5")]
    [InlineData("hebrew2", 73, true, "\u200f\u05ea\u05ea\u05ea\u05d6")]
    [InlineData("hebrew2", 74, true, "\u200f\u05ea\u05ea\u05ea\u05d7")]
    [InlineData("hebrew2", 75, true, "\u200f\u05ea\u05ea\u05ea\u05d8")]
    [InlineData("hebrew2", 76, true, "\u200f\u05ea\u05ea\u05ea\u05d9")]
    [InlineData("hebrew2", 77, true, "\u200f\u05ea\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 78, true, "\u200f\u05ea\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 79, true, "\u200f\u05ea\u05ea\u05ea\u05de")]
    [InlineData("hebrew2", 80, true, "\u200f\u05ea\u05ea\u05ea\u05e0")]
    [InlineData("hebrew2", 81, true, "\u200f\u05ea\u05ea\u05ea\u05e1")]
    [InlineData("hebrew2", 82, true, "\u200f\u05ea\u05ea\u05ea\u05e2")]
    [InlineData("hebrew2", 83, true, "\u200f\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 84, true, "\u200f\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 85, true, "\u200f\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 86, true, "\u200f\u05ea\u05ea\u05ea\u05e8")]
    [InlineData("hebrew2", 87, true, "\u200f\u05ea\u05ea\u05ea\u05e9")]
    [InlineData("hebrew2", 88, true, "\u200f\u05ea\u05ea\u05ea\u05ea")]
    [InlineData("hebrew2", 89, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d0")]
    [InlineData("hebrew2", 90, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d1")]
    [InlineData("hebrew2", 91, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d2")]
    [InlineData("hebrew2", 92, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d3")]
    [InlineData("hebrew2", 93, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 94, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d5")]
    [InlineData("hebrew2", 95, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d6")]
    [InlineData("hebrew2", 96, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d7")]
    [InlineData("hebrew2", 97, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d8")]
    [InlineData("hebrew2", 98, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d9")]
    [InlineData("hebrew2", 99, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 100, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 114, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05d3")]
    [InlineData("hebrew2", 115, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 116, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05d5")]
    [InlineData("hebrew2", 299, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05de")]
    [InlineData("hebrew2", 300, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e0")]
    [InlineData("hebrew2", 301, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e1")]
    [InlineData("hebrew2", 303, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 304, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 305, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 390, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e2")]
    [InlineData("hebrew2", 391, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 392, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 393, true, "\u200f\u05d0")]
    [InlineData("hebrew2", 399, true, "\u200f\u05d6")]
    [InlineData("hebrew2", 400, true, "\u200f\u05d7")]
    [InlineData("hebrew2", 401, true, "\u200f\u05d8")]
    [InlineData("hebrew2", 539, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e1")]
    [InlineData("hebrew2", 540, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e2")]
    [InlineData("hebrew2", 541, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 557, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 558, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 559, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05de")]
    [InlineData("hebrew2", 607, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 608, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 609, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 659, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d2")]
    [InlineData("hebrew2", 660, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d3")]
    [InlineData("hebrew2", 661, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 681, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d2")]
    [InlineData("hebrew2", 682, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d3")]
    [InlineData("hebrew2", 683, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 783, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 784, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 785, true, "\u200f\u05d0")]
    [InlineData("hebrew2", 839, true, "\u200f\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 840, true, "\u200f\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 841, true, "\u200f\u05ea\u05ea\u05de")]
    [InlineData("hebrew2", 867, true, "\u200f\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 868, true, "\u200f\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 869, true, "\u200f\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 999, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4")]
    [InlineData("hebrew2", 1000, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6")]
    [InlineData("hebrew2", 1001, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 1049, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d0")]
    [InlineData("hebrew2", 1050, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d1")]
    [InlineData("hebrew2", 1051, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05d2")]
    [InlineData("hebrew2", 1084, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e0")]
    [InlineData("hebrew2", 1085, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e1")]
    [InlineData("hebrew2", 1086, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e2")]
    [InlineData("hebrew2", 1229, true, "\u200f\u05ea\u05ea\u05d8")]
    [InlineData("hebrew2", 1230, true, "\u200f\u05ea\u05ea\u05d9")]
    [InlineData("hebrew2", 1231, true, "\u200f\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 1270, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d5")]
    [InlineData("hebrew2", 1271, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d6")]
    [InlineData("hebrew2", 1272, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05d7")]
    [InlineData("hebrew2", 2459, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05e7")]
    [InlineData("hebrew2", 2460, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05e8")]
    [InlineData("hebrew2", 2461, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05e9")]
    [InlineData("hebrew2", 32767, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05db")]
    [InlineData("hebrew2", 32768, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05dc")]
    [InlineData("hebrew2", 65535, true, "\u200f\u05ea\u05ea\u05ea\u05d4")]
    [InlineData("hebrew2", 1000000, true, "\u200f\u05d7")]
    [InlineData("hebrew2", 2147483647, true, "\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e7")]
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
