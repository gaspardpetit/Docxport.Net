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

public sealed class TaiwaneseKoreanDigitListMarkerTests
{
    [Theory]
    [InlineData("taiwaneseDigital")]
    [InlineData("koreanDigital")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Word ListString captures with PDF artifacts retained; negative values are start overrides.
    [Theory]
    [InlineData("koreanDigital", -2147483648, true, "")]
    [InlineData("koreanDigital", -1, true, "")]
    [InlineData("koreanDigital", 0, true, "\uc601")]
    [InlineData("koreanDigital", 1, true, "\uc77c")]
    [InlineData("koreanDigital", 2, true, "\uc774")]
    [InlineData("koreanDigital", 9, true, "\uad6c")]
    [InlineData("koreanDigital", 10, true, "\uc77c\uc601")]
    [InlineData("koreanDigital", 21, true, "\uc774\uc77c")]
    [InlineData("koreanDigital", 100, true, "\uc77c\uc601\uc601")]
    [InlineData("koreanDigital", 32766, true, "\uc0bc\uc774\uce60\uc721\uc721")]
    [InlineData("koreanDigital", 32767, true, "\uc0bc\uc774\uce60\uc721\uce60")]
    [InlineData("koreanDigital", 32768, true, "\uc0bc\uc774\uce60\uc721\ud314")]
    [InlineData("koreanDigital", 65534, true, "\uc721\uc624\uc624\uc0bc\uc0ac")]
    [InlineData("koreanDigital", 65535, true, "\uc721\uc624\uc624\uc0bc\uc624")]
    [InlineData("koreanDigital", 999999, true, "\uad6c\uad6c\uad6c\uad6c\uad6c\uad6c")]
    [InlineData("koreanDigital", 1000000, true, "\uc77c\uc601\uc601\uc601\uc601\uc601\uc601")]
    [InlineData("koreanDigital", 2147483647, true, "\uc774\uc77c\uc0ac\uce60\uc0ac\ud314\uc0bc\uc721\uc0ac\uce60")]
    [InlineData("taiwaneseDigital", -2147483648, true, "")]
    [InlineData("taiwaneseDigital", -1, true, "")]
    [InlineData("taiwaneseDigital", 0, true, "\u25cb")]
    [InlineData("taiwaneseDigital", 1, true, "\u4e00")]
    [InlineData("taiwaneseDigital", 2, true, "\u4e8c")]
    [InlineData("taiwaneseDigital", 9, true, "\u4e5d")]
    [InlineData("taiwaneseDigital", 10, true, "\u4e00\u25cb")]
    [InlineData("taiwaneseDigital", 21, true, "\u4e8c\u4e00")]
    [InlineData("taiwaneseDigital", 100, true, "\u4e00\u25cb\u25cb")]
    [InlineData("taiwaneseDigital", 32766, true, "\u4e09\u4e8c\u4e03\u516d\u516d")]
    [InlineData("taiwaneseDigital", 32767, true, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("taiwaneseDigital", 32768, true, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("taiwaneseDigital", 65534, true, "\u516d\u4e94\u4e94\u4e09\u56db")]
    [InlineData("taiwaneseDigital", 65535, true, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("taiwaneseDigital", 999999, true, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseDigital", 1000000, true, "\u4e00\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb")]
    [InlineData("taiwaneseDigital", 2147483647, true, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
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
