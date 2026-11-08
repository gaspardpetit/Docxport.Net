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

public sealed class JapaneseLegalListMarkerTests
{
    [Theory]
    [InlineData("japaneseLegal")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Captured from Word PDFs for standalone start overrides.
    [Theory]
    [InlineData("japaneseLegal", 0, true, "〇")]
    [InlineData("japaneseLegal", 1, true, "壱")]
    [InlineData("japaneseLegal", 10, true, "壱拾")]
    [InlineData("japaneseLegal", 11, true, "壱拾壱")]
    [InlineData("japaneseLegal", 21, true, "弐拾壱")]
    [InlineData("japaneseLegal", 100, true, "壱百")]
    [InlineData("japaneseLegal", 101, true, "壱百壱")]
    [InlineData("japaneseLegal", 110, true, "壱百壱拾")]
    [InlineData("japaneseLegal", 1000, true, "壱阡")]
    [InlineData("japaneseLegal", 1001, true, "壱阡壱")]
    [InlineData("japaneseLegal", 1010, true, "壱阡壱拾")]
    [InlineData("japaneseLegal", 1100, true, "壱阡壱百")]
    [InlineData("japaneseLegal", 10000, true, "壱萬")]
    [InlineData("japaneseLegal", 10001, true, "壱萬壱")]
    [InlineData("japaneseLegal", 10010, true, "壱萬壱拾")]
    [InlineData("japaneseLegal", 10100, true, "壱萬壱百")]
    [InlineData("japaneseLegal", 11000, true, "壱萬壱阡")]
    [InlineData("japaneseLegal", 100000, true, "壱拾萬")]
    [InlineData("japaneseLegal", 100010, true, "壱拾萬壱拾")]
    [InlineData("japaneseLegal", 100100, true, "壱拾萬壱百")]
    [InlineData("japaneseLegal", 101000, true, "壱拾萬壱阡")]
    [InlineData("japaneseLegal", 110000, true, "壱拾壱萬")]
    [InlineData("japaneseLegal", 999999, true, "九拾九萬九阡九百九拾九")]
    [InlineData("japaneseLegal", 1000000, true, "壱百萬")]
    [InlineData("japaneseLegal", 2, true, "\u5f10")]
    [InlineData("japaneseLegal", 3, true, "\u53c2")]
    [InlineData("japaneseLegal", 4, true, "\u56db")]
    [InlineData("japaneseLegal", 5, true, "\u4f0d")]
    [InlineData("japaneseLegal", 6, true, "\u516d")]
    [InlineData("japaneseLegal", 7, true, "\u4e03")]
    [InlineData("japaneseLegal", 8, true, "\u516b")]
    [InlineData("japaneseLegal", 9, true, "\u4e5d")]
    [InlineData("japaneseLegal", 12, true, "\u58f1\u62fe\u5f10")]
    [InlineData("japaneseLegal", 13, true, "\u58f1\u62fe\u53c2")]
    [InlineData("japaneseLegal", 14, true, "\u58f1\u62fe\u56db")]
    [InlineData("japaneseLegal", 15, true, "\u58f1\u62fe\u4f0d")]
    [InlineData("japaneseLegal", 16, true, "\u58f1\u62fe\u516d")]
    [InlineData("japaneseLegal", 17, true, "\u58f1\u62fe\u4e03")]
    [InlineData("japaneseLegal", 18, true, "\u58f1\u62fe\u516b")]
    [InlineData("japaneseLegal", 19, true, "\u58f1\u62fe\u4e5d")]
    [InlineData("japaneseLegal", 20, true, "\u5f10\u62fe")]
    [InlineData("japaneseLegal", 9999999, true, "\u4e5d\u767e\u4e5d\u62fe\u4e5d\u842c\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("japaneseLegal", 100000000, true, "\u58f1\u5104")]
    [InlineData("japaneseLegal", 100000001, true, "\u58f1\u5104\u58f1")]
    [InlineData("japaneseLegal", 100010000, true, "\u58f1\u5104\u58f1\u842c")]
    [InlineData("japaneseLegal", 2147483647, true, "\u5f10\u62fe\u58f1\u5104\u56db\u9621\u4e03\u767e\u56db\u62fe\u516b\u842c\u53c2\u9621\u516d\u767e\u56db\u62fe\u4e03")]
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
