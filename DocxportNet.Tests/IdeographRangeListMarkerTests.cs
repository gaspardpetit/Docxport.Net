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

public sealed class IdeographRangeListMarkerTests
{
    [Theory]
    [InlineData("ideographTraditional")]
    [InlineData("ideographZodiac")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("49", DxpListMarkerFormatter.FormatNumber(49, new NumberFormatValues(format)));
    }

    // Word ListString sequences, corroborated by Word PDF export.
    public static IEnumerable<object[]> WordSequences()
    {
        yield return new object[] { "ideographTraditional", 0, false, new string[] { "0", "\u7532", "\u4e59", "\u4e19", "\u4e01", "\u620a", "\u5df1", "\u5e9a", "\u8f9b", "\u58ec", "\u7678", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60" } };
        yield return new object[] { "ideographZodiac", 0, false, new string[] { "0", "\u5b50", "\u4e11", "\u5bc5", "\u536f", "\u8fb0", "\u5df3", "\u5348", "\u672a", "\u7533", "\u9149", "\u620d", "\u4ea5", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60" } };
    }

    [Theory]
    [MemberData(nameof(WordSequences))]
    public void SuccessiveMarkersMatchWordAcrossIdeographRanges(string format, int start, bool startOverride, string[] expected)
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
