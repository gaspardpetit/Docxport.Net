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

public sealed class ChineseCountingListMarkerTests
{
    [Theory]
    [InlineData("chineseCounting")]
    [InlineData("taiwaneseCounting")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("49", DxpListMarkerFormatter.FormatNumber(49, new NumberFormatValues(format)));
    }

    // Word ListString sequences, corroborated by Word PDF export.
    public static IEnumerable<object[]> WordSequences()
    {
        yield return new object[] { "chineseCounting", 1, false, new string[] { "\u4e00", "\u4e8c", "\u4e09", "\u56db", "\u4e94", "\u516d", "\u4e03", "\u516b", "\u4e5d", "\u5341", "\u5341\u4e00", "\u5341\u4e8c", "\u5341\u4e09", "\u5341\u56db", "\u5341\u4e94", "\u5341\u516d", "\u5341\u4e03", "\u5341\u516b", "\u5341\u4e5d", "\u4e8c\u5341", "\u4e8c\u5341\u4e00" } };
        yield return new object[] { "chineseCounting", 98, false, new string[] { "\u4e5d\u5341\u516b", "\u4e5d\u5341\u4e5d", "\u4e00\u25cb\u25cb", "\u4e00\u25cb\u4e00", "\u4e00\u25cb\u4e8c" } };
        yield return new object[] { "chineseCounting", 998, false, new string[] { "\u4e5d\u4e5d\u516b", "\u4e5d\u4e5d\u4e5d", "\u4e00\u25cb\u25cb\u25cb", "\u4e00\u25cb\u25cb\u4e00", "\u4e00\u25cb\u25cb\u4e8c" } };
        yield return new object[] { "chineseCounting", 9998, false, new string[] { "\u4e5d\u4e5d\u4e5d\u516b", "\u4e5d\u4e5d\u4e5d\u4e5d", "\u4e00\u25cb\u25cb\u25cb\u25cb", "\u4e00\u25cb\u25cb\u25cb\u4e00", "\u4e00\u25cb\u25cb\u25cb\u4e8c" } };
        yield return new object[] { "taiwaneseCounting", 1, false, new string[] { "\u4e00", "\u4e8c", "\u4e09", "\u56db", "\u4e94", "\u516d", "\u4e03", "\u516b", "\u4e5d", "\u5341", "\u5341\u4e00", "\u5341\u4e8c", "\u5341\u4e09", "\u5341\u56db", "\u5341\u4e94", "\u5341\u516d", "\u5341\u4e03", "\u5341\u516b", "\u5341\u4e5d", "\u4e8c\u5341", "\u4e8c\u5341\u4e00" } };
        yield return new object[] { "taiwaneseCounting", 98, false, new string[] { "\u4e5d\u5341\u516b", "\u4e5d\u5341\u4e5d", "\u4e00\u25cb\u25cb", "\u4e00\u25cb\u4e00", "\u4e00\u25cb\u4e8c" } };
        yield return new object[] { "taiwaneseCounting", 998, false, new string[] { "\u4e5d\u4e5d\u516b", "\u4e5d\u4e5d\u4e5d", "\u4e00\u25cb\u25cb\u25cb", "\u4e00\u25cb\u25cb\u4e00", "\u4e00\u25cb\u25cb\u4e8c" } };
        yield return new object[] { "taiwaneseCounting", 9998, false, new string[] { "\u4e5d\u4e5d\u4e5d\u516b", "\u4e5d\u4e5d\u4e5d\u4e5d", "\u4e00\u25cb\u25cb\u25cb\u25cb", "\u4e00\u25cb\u25cb\u25cb\u4e00", "\u4e00\u25cb\u25cb\u25cb\u4e8c" } };
    }

    [Theory]
    [MemberData(nameof(WordSequences))]
    public void SuccessiveMarkersMatchWordAcrossChineseCountingBoundaries(string format, int start, bool startOverride, string[] expected)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var instance = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 };
        if (startOverride)
            instance.Append(new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = 0 });
        var numbering = new Numbering(new AbstractNum(
            new Level(new StartNumberingValue { Val = startOverride ? 1 : start },
                new NumberingFormat { Val = new NumberFormatValues(format) },
                new LevelText { Val = "BEGIN%1END" }) { LevelIndex = 0 }) { AbstractNumberId = 1 }, instance);
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        main.Document = new Document(new Body(Enumerable.Range(0, expected.Length).Select(i =>
            new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
                new Run(new Text($"Item{i}"))))));
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            Assert.Equal(expected, Regex.Matches(output, "BEGIN(.*?)END").Select(m => m.Groups[1].Value));
            Assert.Equal(expected.Length, Regex.Matches(output, @"Item\d+").Count);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
