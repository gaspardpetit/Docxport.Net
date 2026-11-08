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

public sealed class ZodiacCycleListMarkerTests
{
    [Fact]
    public void SharedFieldFormatterRetainsDecimalFallback()
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, NumberFormatValues.IdeographZodiacTraditional));
        Assert.Equal("61", DxpListMarkerFormatter.FormatNumber(61, NumberFormatValues.IdeographZodiacTraditional));
    }

    // Captured from Word ListString; includes complete symbol cycles and counter boundaries.
    [Theory]
    [InlineData("ideographZodiacTraditional", -2147483648, true, "")]
    [InlineData("ideographZodiacTraditional", -1, true, "-1")]
    [InlineData("ideographZodiacTraditional", 0, true, "0")]
    [InlineData("ideographZodiacTraditional", 1, true, "\u7532\u5b50")]
    [InlineData("ideographZodiacTraditional", 2, true, "\u4e59\u4e11")]
    [InlineData("ideographZodiacTraditional", 3, true, "\u4e19\u5bc5")]
    [InlineData("ideographZodiacTraditional", 4, true, "\u4e01\u536f")]
    [InlineData("ideographZodiacTraditional", 5, true, "\u620a\u8fb0")]
    [InlineData("ideographZodiacTraditional", 6, true, "\u5df1\u5df3")]
    [InlineData("ideographZodiacTraditional", 7, true, "\u5e9a\u5348")]
    [InlineData("ideographZodiacTraditional", 8, true, "\u8f9b\u672a")]
    [InlineData("ideographZodiacTraditional", 9, true, "\u58ec\u7533")]
    [InlineData("ideographZodiacTraditional", 10, true, "\u7678\u9149")]
    [InlineData("ideographZodiacTraditional", 11, true, "\u7532\u620d")]
    [InlineData("ideographZodiacTraditional", 12, true, "\u4e59\u4ea5")]
    [InlineData("ideographZodiacTraditional", 13, true, "\u4e19\u5b50")]
    [InlineData("ideographZodiacTraditional", 14, true, "\u4e01\u4e11")]
    [InlineData("ideographZodiacTraditional", 15, true, "\u620a\u5bc5")]
    [InlineData("ideographZodiacTraditional", 16, true, "\u5df1\u536f")]
    [InlineData("ideographZodiacTraditional", 17, true, "\u5e9a\u8fb0")]
    [InlineData("ideographZodiacTraditional", 18, true, "\u8f9b\u5df3")]
    [InlineData("ideographZodiacTraditional", 19, true, "\u58ec\u5348")]
    [InlineData("ideographZodiacTraditional", 20, true, "\u7678\u672a")]
    [InlineData("ideographZodiacTraditional", 21, true, "\u7532\u7533")]
    [InlineData("ideographZodiacTraditional", 22, true, "\u4e59\u9149")]
    [InlineData("ideographZodiacTraditional", 23, true, "\u4e19\u620d")]
    [InlineData("ideographZodiacTraditional", 24, true, "\u4e01\u4ea5")]
    [InlineData("ideographZodiacTraditional", 25, true, "\u620a\u5b50")]
    [InlineData("ideographZodiacTraditional", 26, true, "\u5df1\u4e11")]
    [InlineData("ideographZodiacTraditional", 27, true, "\u5e9a\u5bc5")]
    [InlineData("ideographZodiacTraditional", 28, true, "\u8f9b\u536f")]
    [InlineData("ideographZodiacTraditional", 29, true, "\u58ec\u8fb0")]
    [InlineData("ideographZodiacTraditional", 30, true, "\u7678\u5df3")]
    [InlineData("ideographZodiacTraditional", 31, true, "\u7532\u5348")]
    [InlineData("ideographZodiacTraditional", 32, true, "\u4e59\u672a")]
    [InlineData("ideographZodiacTraditional", 33, true, "\u4e19\u7533")]
    [InlineData("ideographZodiacTraditional", 34, true, "\u4e01\u9149")]
    [InlineData("ideographZodiacTraditional", 35, true, "\u620a\u620d")]
    [InlineData("ideographZodiacTraditional", 36, true, "\u5df1\u4ea5")]
    [InlineData("ideographZodiacTraditional", 37, true, "\u5e9a\u5b50")]
    [InlineData("ideographZodiacTraditional", 38, true, "\u8f9b\u4e11")]
    [InlineData("ideographZodiacTraditional", 39, true, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 40, true, "\u7678\u536f")]
    [InlineData("ideographZodiacTraditional", 41, true, "\u7532\u8fb0")]
    [InlineData("ideographZodiacTraditional", 42, true, "\u4e59\u5df3")]
    [InlineData("ideographZodiacTraditional", 43, true, "\u4e19\u5348")]
    [InlineData("ideographZodiacTraditional", 44, true, "\u4e01\u672a")]
    [InlineData("ideographZodiacTraditional", 45, true, "\u620a\u7533")]
    [InlineData("ideographZodiacTraditional", 46, true, "\u5df1\u9149")]
    [InlineData("ideographZodiacTraditional", 47, true, "\u5e9a\u620d")]
    [InlineData("ideographZodiacTraditional", 48, true, "\u8f9b\u4ea5")]
    [InlineData("ideographZodiacTraditional", 49, true, "\u58ec\u5b50")]
    [InlineData("ideographZodiacTraditional", 50, true, "\u7678\u4e11")]
    [InlineData("ideographZodiacTraditional", 51, true, "\u7532\u5bc5")]
    [InlineData("ideographZodiacTraditional", 52, true, "\u4e59\u536f")]
    [InlineData("ideographZodiacTraditional", 53, true, "\u4e19\u8fb0")]
    [InlineData("ideographZodiacTraditional", 54, true, "\u4e01\u5df3")]
    [InlineData("ideographZodiacTraditional", 55, true, "\u620a\u5348")]
    [InlineData("ideographZodiacTraditional", 56, true, "\u5df1\u672a")]
    [InlineData("ideographZodiacTraditional", 57, true, "\u5e9a\u7533")]
    [InlineData("ideographZodiacTraditional", 58, true, "\u8f9b\u9149")]
    [InlineData("ideographZodiacTraditional", 59, true, "\u58ec\u620d")]
    [InlineData("ideographZodiacTraditional", 60, true, "\u7678\u4ea5")]
    [InlineData("ideographZodiacTraditional", 61, true, "\u7532\u5b50")]
    [InlineData("ideographZodiacTraditional", 62, true, "\u4e59\u4e11")]
    [InlineData("ideographZodiacTraditional", 63, true, "\u4e19\u5bc5")]
    [InlineData("ideographZodiacTraditional", 64, true, "\u4e01\u536f")]
    [InlineData("ideographZodiacTraditional", 65, true, "\u620a\u8fb0")]
    [InlineData("ideographZodiacTraditional", 66, true, "\u5df1\u5df3")]
    [InlineData("ideographZodiacTraditional", 67, true, "\u5e9a\u5348")]
    [InlineData("ideographZodiacTraditional", 68, true, "\u8f9b\u672a")]
    [InlineData("ideographZodiacTraditional", 69, true, "\u58ec\u7533")]
    [InlineData("ideographZodiacTraditional", 70, true, "\u7678\u9149")]
    [InlineData("ideographZodiacTraditional", 71, true, "\u7532\u620d")]
    [InlineData("ideographZodiacTraditional", 72, true, "\u4e59\u4ea5")]
    [InlineData("ideographZodiacTraditional", 73, true, "\u4e19\u5b50")]
    [InlineData("ideographZodiacTraditional", 74, true, "\u4e01\u4e11")]
    [InlineData("ideographZodiacTraditional", 75, true, "\u620a\u5bc5")]
    [InlineData("ideographZodiacTraditional", 76, true, "\u5df1\u536f")]
    [InlineData("ideographZodiacTraditional", 77, true, "\u5e9a\u8fb0")]
    [InlineData("ideographZodiacTraditional", 78, true, "\u8f9b\u5df3")]
    [InlineData("ideographZodiacTraditional", 79, true, "\u58ec\u5348")]
    [InlineData("ideographZodiacTraditional", 80, true, "\u7678\u672a")]
    [InlineData("ideographZodiacTraditional", 81, true, "\u7532\u7533")]
    [InlineData("ideographZodiacTraditional", 82, true, "\u4e59\u9149")]
    [InlineData("ideographZodiacTraditional", 83, true, "\u4e19\u620d")]
    [InlineData("ideographZodiacTraditional", 84, true, "\u4e01\u4ea5")]
    [InlineData("ideographZodiacTraditional", 85, true, "\u620a\u5b50")]
    [InlineData("ideographZodiacTraditional", 86, true, "\u5df1\u4e11")]
    [InlineData("ideographZodiacTraditional", 87, true, "\u5e9a\u5bc5")]
    [InlineData("ideographZodiacTraditional", 88, true, "\u8f9b\u536f")]
    [InlineData("ideographZodiacTraditional", 89, true, "\u58ec\u8fb0")]
    [InlineData("ideographZodiacTraditional", 90, true, "\u7678\u5df3")]
    [InlineData("ideographZodiacTraditional", 91, true, "\u7532\u5348")]
    [InlineData("ideographZodiacTraditional", 92, true, "\u4e59\u672a")]
    [InlineData("ideographZodiacTraditional", 93, true, "\u4e19\u7533")]
    [InlineData("ideographZodiacTraditional", 94, true, "\u4e01\u9149")]
    [InlineData("ideographZodiacTraditional", 95, true, "\u620a\u620d")]
    [InlineData("ideographZodiacTraditional", 96, true, "\u5df1\u4ea5")]
    [InlineData("ideographZodiacTraditional", 97, true, "\u5e9a\u5b50")]
    [InlineData("ideographZodiacTraditional", 98, true, "\u8f9b\u4e11")]
    [InlineData("ideographZodiacTraditional", 99, true, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 100, true, "\u7678\u536f")]
    [InlineData("ideographZodiacTraditional", 101, true, "\u7532\u8fb0")]
    [InlineData("ideographZodiacTraditional", 102, true, "\u4e59\u5df3")]
    [InlineData("ideographZodiacTraditional", 103, true, "\u4e19\u5348")]
    [InlineData("ideographZodiacTraditional", 104, true, "\u4e01\u672a")]
    [InlineData("ideographZodiacTraditional", 105, true, "\u620a\u7533")]
    [InlineData("ideographZodiacTraditional", 106, true, "\u5df1\u9149")]
    [InlineData("ideographZodiacTraditional", 107, true, "\u5e9a\u620d")]
    [InlineData("ideographZodiacTraditional", 108, true, "\u8f9b\u4ea5")]
    [InlineData("ideographZodiacTraditional", 109, true, "\u58ec\u5b50")]
    [InlineData("ideographZodiacTraditional", 110, true, "\u7678\u4e11")]
    [InlineData("ideographZodiacTraditional", 111, true, "\u7532\u5bc5")]
    [InlineData("ideographZodiacTraditional", 112, true, "\u4e59\u536f")]
    [InlineData("ideographZodiacTraditional", 113, true, "\u4e19\u8fb0")]
    [InlineData("ideographZodiacTraditional", 114, true, "\u4e01\u5df3")]
    [InlineData("ideographZodiacTraditional", 115, true, "\u620a\u5348")]
    [InlineData("ideographZodiacTraditional", 116, true, "\u5df1\u672a")]
    [InlineData("ideographZodiacTraditional", 117, true, "\u5e9a\u7533")]
    [InlineData("ideographZodiacTraditional", 118, true, "\u8f9b\u9149")]
    [InlineData("ideographZodiacTraditional", 119, true, "\u58ec\u620d")]
    [InlineData("ideographZodiacTraditional", 120, true, "\u7678\u4ea5")]
    [InlineData("ideographZodiacTraditional", 121, true, "\u7532\u5b50")]
    [InlineData("ideographZodiacTraditional", 420, true, "\u7678\u4ea5")]
    [InlineData("ideographZodiacTraditional", 421, true, "\u7532\u5b50")]
    [InlineData("ideographZodiacTraditional", 32767, true, "\u5e9a\u5348")]
    [InlineData("ideographZodiacTraditional", 32768, true, "\u8f9b\u672a")]
    [InlineData("ideographZodiacTraditional", 65535, true, "\u620a\u5bc5")]
    [InlineData("ideographZodiacTraditional", 999999, true, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 1000000, true, "\u7678\u536f")]
    [InlineData("ideographZodiacTraditional", 2147483647, true, "\u5e9a\u5348")]
    public void RegionalSymbolsMatchWordAtBoundaries(string format, int start, bool startOverride, string? expected)
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
