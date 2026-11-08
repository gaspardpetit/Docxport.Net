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

public sealed class KoreanUnitListMarkerTests
{
    [Theory]
    [InlineData("koreanCounting")]
    [InlineData("koreanLegal")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Captured from Word PDFs for standalone start overrides.
    [Theory]
    [InlineData("koreanCounting", 0, true, "\uc601")]
    [InlineData("koreanCounting", 1, true, "\uc77c")]
    [InlineData("koreanCounting", 10, true, "\uc2ed")]
    [InlineData("koreanCounting", 11, true, "\uc2ed\uc77c")]
    [InlineData("koreanCounting", 21, true, "\uc774\uc2ed\uc77c")]
    [InlineData("koreanCounting", 100, true, "\ubc31")]
    [InlineData("koreanCounting", 101, true, "\ubc31\uc77c")]
    [InlineData("koreanCounting", 110, true, "\ubc31\uc2ed")]
    [InlineData("koreanCounting", 1000, true, "\ucc9c")]
    [InlineData("koreanCounting", 1001, true, "\ucc9c\uc77c")]
    [InlineData("koreanCounting", 1010, true, "\ucc9c\uc2ed")]
    [InlineData("koreanCounting", 1100, true, "\ucc9c\ubc31")]
    [InlineData("koreanCounting", 10000, true, "\ub9cc")]
    [InlineData("koreanCounting", 10001, true, "\ub9cc\uc77c")]
    [InlineData("koreanCounting", 10010, true, "\ub9cc\uc2ed")]
    [InlineData("koreanCounting", 10100, true, "\ub9cc\ubc31")]
    [InlineData("koreanCounting", 11000, true, "\ub9cc\uc77c\ucc9c")]
    [InlineData("koreanCounting", 100000, true, "\uc2ed\ub9cc")]
    [InlineData("koreanCounting", 100010, true, "\uc2ed\ub9cc\uc2ed")]
    [InlineData("koreanCounting", 100100, true, "\uc2ed\ub9cc\ubc31")]
    [InlineData("koreanCounting", 101000, true, "\uc2ed\ub9cc\uc77c\ucc9c")]
    [InlineData("koreanCounting", 110000, true, "\uc2ed\ub9cc")]
    [InlineData("koreanCounting", 999999, true, "\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 1000000, true, "")]
    [InlineData("koreanCounting", 2, true, "\uc774")]
    [InlineData("koreanCounting", 3, true, "\uc0bc")]
    [InlineData("koreanCounting", 4, true, "\uc0ac")]
    [InlineData("koreanCounting", 5, true, "\uc624")]
    [InlineData("koreanCounting", 6, true, "\uc721")]
    [InlineData("koreanCounting", 7, true, "\uce60")]
    [InlineData("koreanCounting", 8, true, "\ud314")]
    [InlineData("koreanCounting", 9, true, "\uad6c")]
    [InlineData("koreanCounting", 12, true, "\uc2ed\uc774")]
    [InlineData("koreanCounting", 13, true, "\uc2ed\uc0bc")]
    [InlineData("koreanCounting", 14, true, "\uc2ed\uc0ac")]
    [InlineData("koreanCounting", 15, true, "\uc2ed\uc624")]
    [InlineData("koreanCounting", 16, true, "\uc2ed\uc721")]
    [InlineData("koreanCounting", 17, true, "\uc2ed\uce60")]
    [InlineData("koreanCounting", 18, true, "\uc2ed\ud314")]
    [InlineData("koreanCounting", 19, true, "\uc2ed\uad6c")]
    [InlineData("koreanCounting", 20, true, "\uc774\uc2ed")]
    [InlineData("koreanCounting", 22, true, "\uc774\uc2ed\uc774")]
    [InlineData("koreanCounting", 23, true, "\uc774\uc2ed\uc0bc")]
    [InlineData("koreanCounting", 24, true, "\uc774\uc2ed\uc0ac")]
    [InlineData("koreanCounting", 25, true, "\uc774\uc2ed\uc624")]
    [InlineData("koreanCounting", 26, true, "\uc774\uc2ed\uc721")]
    [InlineData("koreanCounting", 27, true, "\uc774\uc2ed\uce60")]
    [InlineData("koreanCounting", 28, true, "\uc774\uc2ed\ud314")]
    [InlineData("koreanCounting", 29, true, "\uc774\uc2ed\uad6c")]
    [InlineData("koreanCounting", 30, true, "\uc0bc\uc2ed")]
    [InlineData("koreanCounting", 31, true, "\uc0bc\uc2ed\uc77c")]
    [InlineData("koreanCounting", 32, true, "\uc0bc\uc2ed\uc774")]
    [InlineData("koreanCounting", 33, true, "\uc0bc\uc2ed\uc0bc")]
    [InlineData("koreanCounting", 34, true, "\uc0bc\uc2ed\uc0ac")]
    [InlineData("koreanCounting", 35, true, "\uc0bc\uc2ed\uc624")]
    [InlineData("koreanCounting", 36, true, "\uc0bc\uc2ed\uc721")]
    [InlineData("koreanCounting", 37, true, "\uc0bc\uc2ed\uce60")]
    [InlineData("koreanCounting", 38, true, "\uc0bc\uc2ed\ud314")]
    [InlineData("koreanCounting", 39, true, "\uc0bc\uc2ed\uad6c")]
    [InlineData("koreanCounting", 40, true, "\uc0ac\uc2ed")]
    [InlineData("koreanCounting", 99, true, "\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 111, true, "\ubc31\uc2ed\uc77c")]
    [InlineData("koreanCounting", 999, true, "\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 1111, true, "\ucc9c\ubc31\uc2ed\uc77c")]
    [InlineData("koreanCounting", 20000, true, "\uc774\ub9cc")]
    [InlineData("koreanCounting", 111000, true, "\uc2ed\ub9cc\uc77c\ucc9c")]
    [InlineData("koreanCounting", 111100, true, "\uc2ed\ub9cc\uc77c\ucc9c\ubc31")]
    [InlineData("koreanCounting", 111110, true, "\uc2ed\ub9cc\uc77c\ucc9c\ubc31\uc2ed")]
    [InlineData("koreanCounting", 111111, true, "\uc2ed\ub9cc\uc77c\ucc9c\ubc31\uc2ed\uc77c")]
    [InlineData("koreanCounting", 120000, true, "\uc2ed\uc774\ub9cc")]
    [InlineData("koreanCounting", 210000, true, "\uc774\uc2ed\ub9cc")]
    [InlineData("koreanCounting", 10000000, true, "")]
    [InlineData("koreanCounting", 100000000, true, "")]
    [InlineData("koreanCounting", 2147483647, true, "")]
    [InlineData("koreanCounting", 50, true, "\uc624\uc2ed")]
    [InlineData("koreanCounting", 60, true, "\uc721\uc2ed")]
    [InlineData("koreanCounting", 70, true, "\uce60\uc2ed")]
    [InlineData("koreanCounting", 80, true, "\ud314\uc2ed")]
    [InlineData("koreanCounting", 90, true, "\uad6c\uc2ed")]
    [InlineData("koreanCounting", 1000001, true, "")]
    [InlineData("koreanCounting", 1100000, true, "")]
    [InlineData("koreanCounting", 1110000, true, "")]
    [InlineData("koreanCounting", 2000000, true, "")]
    [InlineData("koreanCounting", 9999998, true, "")]
    [InlineData("koreanCounting", 9999999, true, "")]
    [InlineData("koreanLegal", 0, true, "0")]
    [InlineData("koreanLegal", 1, true, "\ud558\ub098")]
    [InlineData("koreanLegal", 10, true, "\uc5f4")]
    [InlineData("koreanLegal", 11, true, "\uc5f4\ud558\ub098")]
    [InlineData("koreanLegal", 21, true, "\uc2a4\ubb3c\ud558\ub098")]
    [InlineData("koreanLegal", 100, true, "\ubc31")]
    [InlineData("koreanLegal", 101, true, "\ubc31\uc77c")]
    [InlineData("koreanLegal", 110, true, "\ubc31\uc2ed")]
    [InlineData("koreanLegal", 1000, true, "\ucc9c")]
    [InlineData("koreanLegal", 1001, true, "\ucc9c\uc77c")]
    [InlineData("koreanLegal", 1010, true, "\ucc9c\uc2ed")]
    [InlineData("koreanLegal", 1100, true, "\ucc9c\ubc31")]
    [InlineData("koreanLegal", 10000, true, "\ub9cc")]
    [InlineData("koreanLegal", 10001, true, "\ub9cc\uc77c")]
    [InlineData("koreanLegal", 10010, true, "\ub9cc\uc2ed")]
    [InlineData("koreanLegal", 10100, true, "\ub9cc\ubc31")]
    [InlineData("koreanLegal", 11000, true, "\ub9cc\uc77c\ucc9c")]
    [InlineData("koreanLegal", 100000, true, "\uc2ed\ub9cc")]
    [InlineData("koreanLegal", 100010, true, "\uc2ed\ub9cc\uc2ed")]
    [InlineData("koreanLegal", 100100, true, "\uc2ed\ub9cc\ubc31")]
    [InlineData("koreanLegal", 101000, true, "\uc2ed\ub9cc\uc77c\ucc9c")]
    [InlineData("koreanLegal", 110000, true, "\uc2ed\ub9cc")]
    [InlineData("koreanLegal", 999999, true, "\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanLegal", 1000000, true, "\ubc31\ub9cc")]
    [InlineData("koreanLegal", 2, true, "\ub458")]
    [InlineData("koreanLegal", 3, true, "\uc14b")]
    [InlineData("koreanLegal", 4, true, "\ub137")]
    [InlineData("koreanLegal", 5, true, "\ub2e4\uc12f")]
    [InlineData("koreanLegal", 6, true, "\uc5ec\uc12f")]
    [InlineData("koreanLegal", 7, true, "\uc77c\uacf1")]
    [InlineData("koreanLegal", 8, true, "\uc5ec\ub35f")]
    [InlineData("koreanLegal", 9, true, "\uc544\ud649")]
    [InlineData("koreanLegal", 12, true, "\uc5f4\ub458")]
    [InlineData("koreanLegal", 13, true, "\uc5f4\uc14b")]
    [InlineData("koreanLegal", 14, true, "\uc5f4\ub137")]
    [InlineData("koreanLegal", 15, true, "\uc5f4\ub2e4\uc12f")]
    [InlineData("koreanLegal", 16, true, "\uc5f4\uc5ec\uc12f")]
    [InlineData("koreanLegal", 17, true, "\uc5f4\uc77c\uacf1")]
    [InlineData("koreanLegal", 18, true, "\uc5f4\uc5ec\ub35f")]
    [InlineData("koreanLegal", 19, true, "\uc5f4\uc544\ud649")]
    [InlineData("koreanLegal", 20, true, "\uc2a4\ubb3c")]
    [InlineData("koreanLegal", 22, true, "\uc2a4\ubb3c\ub458")]
    [InlineData("koreanLegal", 23, true, "\uc2a4\ubb3c\uc14b")]
    [InlineData("koreanLegal", 24, true, "\uc2a4\ubb3c\ub137")]
    [InlineData("koreanLegal", 25, true, "\uc2a4\ubb3c\ub2e4\uc12f")]
    [InlineData("koreanLegal", 26, true, "\uc2a4\ubb3c\uc5ec\uc12f")]
    [InlineData("koreanLegal", 27, true, "\uc2a4\ubb3c\uc77c\uacf1")]
    [InlineData("koreanLegal", 28, true, "\uc2a4\ubb3c\uc5ec\ub35f")]
    [InlineData("koreanLegal", 29, true, "\uc2a4\ubb3c\uc544\ud649")]
    [InlineData("koreanLegal", 30, true, "\uc11c\ub978")]
    [InlineData("koreanLegal", 31, true, "\uc11c\ub978\ud558\ub098")]
    [InlineData("koreanLegal", 32, true, "\uc11c\ub978\ub458")]
    [InlineData("koreanLegal", 33, true, "\uc11c\ub978\uc14b")]
    [InlineData("koreanLegal", 34, true, "\uc11c\ub978\ub137")]
    [InlineData("koreanLegal", 35, true, "\uc11c\ub978\ub2e4\uc12f")]
    [InlineData("koreanLegal", 36, true, "\uc11c\ub978\uc5ec\uc12f")]
    [InlineData("koreanLegal", 37, true, "\uc11c\ub978\uc77c\uacf1")]
    [InlineData("koreanLegal", 38, true, "\uc11c\ub978\uc5ec\ub35f")]
    [InlineData("koreanLegal", 39, true, "\uc11c\ub978\uc544\ud649")]
    [InlineData("koreanLegal", 40, true, "\ub9c8\ud754")]
    [InlineData("koreanLegal", 99, true, "\uc544\ud754\uc544\ud649")]
    [InlineData("koreanLegal", 111, true, "\ubc31\uc2ed\uc77c")]
    [InlineData("koreanLegal", 999, true, "\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanLegal", 1111, true, "\ucc9c\ubc31\uc2ed\uc77c")]
    [InlineData("koreanLegal", 20000, true, "\uc774\ub9cc")]
    [InlineData("koreanLegal", 111000, true, "\uc2ed\ub9cc\uc77c\ucc9c")]
    [InlineData("koreanLegal", 111100, true, "\uc2ed\ub9cc\uc77c\ucc9c\ubc31")]
    [InlineData("koreanLegal", 111110, true, "\uc2ed\ub9cc\uc77c\ucc9c\ubc31\uc2ed")]
    [InlineData("koreanLegal", 111111, true, "\uc2ed\ub9cc\uc77c\ucc9c\ubc31\uc2ed\uc77c")]
    [InlineData("koreanLegal", 120000, true, "\uc2ed\uc774\ub9cc")]
    [InlineData("koreanLegal", 210000, true, "\uc774\uc2ed\ub9cc")]
    [InlineData("koreanLegal", 10000000, true, "")]
    [InlineData("koreanLegal", 100000000, true, "")]
    [InlineData("koreanLegal", 2147483647, true, "")]
    [InlineData("koreanLegal", 50, true, "\uc270")]
    [InlineData("koreanLegal", 60, true, "\uc608\uc21c")]
    [InlineData("koreanLegal", 70, true, "\uc77c\ud754")]
    [InlineData("koreanLegal", 80, true, "\uc5ec\ub4e0")]
    [InlineData("koreanLegal", 90, true, "\uc544\ud754")]
    [InlineData("koreanLegal", 1000001, true, "\ubc31\ub9cc\uc77c")]
    [InlineData("koreanLegal", 1100000, true, "\ubc31\uc77c\uc2ed\ub9cc")]
    [InlineData("koreanLegal", 1110000, true, "\ubc31\uc77c\uc2ed\ub9cc")]
    [InlineData("koreanLegal", 2000000, true, "\uc774\ubc31\ub9cc")]
    [InlineData("koreanLegal", 9999998, true, "\uad6c\ubc31\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\ud314")]
    [InlineData("koreanLegal", 9999999, true, "\uad6c\ubc31\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
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
