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

public sealed class ChineseLegalListMarkerTests
{
    [Theory]
    [InlineData("chineseLegalSimplified")]
    [InlineData("ideographLegalTraditional")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Captured from Word PDFs for standalone start overrides.
    [Theory]
    [InlineData("chineseLegalSimplified", 0, true, "\u96f6")]
    [InlineData("chineseLegalSimplified", 1, true, "\u58f9")]
    [InlineData("chineseLegalSimplified", 10, true, "\u58f9\u62fe")]
    [InlineData("chineseLegalSimplified", 11, true, "\u58f9\u62fe\u58f9")]
    [InlineData("chineseLegalSimplified", 21, true, "\u8d30\u62fe\u58f9")]
    [InlineData("chineseLegalSimplified", 100, true, "\u58f9\u4f70")]
    [InlineData("chineseLegalSimplified", 101, true, "\u58f9\u4f70\u96f6\u58f9")]
    [InlineData("chineseLegalSimplified", 110, true, "\u58f9\u4f70\u58f9\u62fe")]
    [InlineData("chineseLegalSimplified", 1000, true, "\u58f9\u4edf")]
    [InlineData("chineseLegalSimplified", 1001, true, "\u58f9\u4edf\u96f6\u58f9")]
    [InlineData("chineseLegalSimplified", 1010, true, "\u58f9\u4edf\u96f6\u58f9\u62fe")]
    [InlineData("chineseLegalSimplified", 1100, true, "\u58f9\u4edf\u58f9\u4f70")]
    [InlineData("chineseLegalSimplified", 10000, true, "\u58f9\u842c")]
    [InlineData("chineseLegalSimplified", 10001, true, "\u58f9\u842c\u96f6\u58f9")]
    [InlineData("chineseLegalSimplified", 10010, true, "\u58f9\u842c\u96f6\u58f9\u62fe")]
    [InlineData("chineseLegalSimplified", 10100, true, "\u58f9\u842c\u96f6\u58f9\u4f70")]
    [InlineData("chineseLegalSimplified", 11000, true, "\u58f9\u842c\u58f9\u4edf")]
    [InlineData("chineseLegalSimplified", 100000, true, "\u58f9\u62fe\u842c")]
    [InlineData("chineseLegalSimplified", 100010, true, "\u58f9\u62fe\u842c\u96f6\u58f9\u62fe")]
    [InlineData("chineseLegalSimplified", 100100, true, "\u58f9\u62fe\u842c\u96f6\u58f9\u4f70")]
    [InlineData("chineseLegalSimplified", 101000, true, "\u58f9\u62fe\u842c\u96f6\u58f9\u4edf")]
    [InlineData("chineseLegalSimplified", 110000, true, "\u58f9\u62fe\u58f9\u842c")]
    [InlineData("chineseLegalSimplified", 999999, true, "\u7396\u62fe\u7396\u842c\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("chineseLegalSimplified", 1000000, true, "")]
    [InlineData("ideographLegalTraditional", 0, true, "\u96f6")]
    [InlineData("ideographLegalTraditional", 1, true, "\u58f9")]
    [InlineData("ideographLegalTraditional", 10, true, "\u58f9\u62fe")]
    [InlineData("ideographLegalTraditional", 11, true, "\u58f9\u62fe\u58f9")]
    [InlineData("ideographLegalTraditional", 21, true, "\u8cb3\u62fe\u58f9")]
    [InlineData("ideographLegalTraditional", 100, true, "\u58f9\u4f70")]
    [InlineData("ideographLegalTraditional", 101, true, "\u58f9\u4f70\u96f6\u58f9")]
    [InlineData("ideographLegalTraditional", 110, true, "\u58f9\u4f70\u58f9\u62fe")]
    [InlineData("ideographLegalTraditional", 1000, true, "\u58f9\u4edf")]
    [InlineData("ideographLegalTraditional", 1001, true, "\u58f9\u4edf\u96f6\u58f9")]
    [InlineData("ideographLegalTraditional", 1010, true, "\u58f9\u4edf\u96f6\u58f9\u62fe")]
    [InlineData("ideographLegalTraditional", 1100, true, "\u58f9\u4edf\u58f9\u4f70")]
    [InlineData("ideographLegalTraditional", 10000, true, "\u58f9\u842c")]
    [InlineData("ideographLegalTraditional", 10001, true, "\u58f9\u842c\u96f6\u58f9")]
    [InlineData("ideographLegalTraditional", 10010, true, "\u58f9\u842c\u96f6\u58f9\u62fe")]
    [InlineData("ideographLegalTraditional", 10100, true, "\u58f9\u842c\u96f6\u58f9\u4f70")]
    [InlineData("ideographLegalTraditional", 11000, true, "\u58f9\u842c\u58f9\u4edf")]
    [InlineData("ideographLegalTraditional", 100000, true, "\u58f9\u62fe\u842c")]
    [InlineData("ideographLegalTraditional", 100010, true, "\u58f9\u62fe\u842c\u96f6\u58f9\u62fe")]
    [InlineData("ideographLegalTraditional", 100100, true, "\u58f9\u62fe\u842c\u96f6\u58f9\u4f70")]
    [InlineData("ideographLegalTraditional", 101000, true, "\u58f9\u62fe\u842c\u96f6\u58f9\u4edf")]
    [InlineData("ideographLegalTraditional", 110000, true, "\u58f9\u62fe\u58f9\u842c")]
    [InlineData("ideographLegalTraditional", 999999, true, "\u7396\u62fe\u7396\u842c\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("ideographLegalTraditional", 1000000, true, "")]
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
