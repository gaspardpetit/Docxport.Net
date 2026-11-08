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

public sealed class EnclosedIdeographListMarkerTests
{
    [Fact]
    public void SharedFieldFormatterRetainsDecimalFallback()
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, NumberFormatValues.IdeographEnclosedCircle));
        Assert.Equal("11", DxpListMarkerFormatter.FormatNumber(11, NumberFormatValues.IdeographEnclosedCircle));
    }

    // Captured from Word ListString; includes complete symbol cycles and counter boundaries.
    [Theory]
    [InlineData("ideographEnclosedCircle", -2147483648, true, "")]
    [InlineData("ideographEnclosedCircle", -1, true, "-1")]
    [InlineData("ideographEnclosedCircle", 0, true, "0")]
    [InlineData("ideographEnclosedCircle", 1, true, "\u3220")]
    [InlineData("ideographEnclosedCircle", 2, true, "\u3221")]
    [InlineData("ideographEnclosedCircle", 3, true, "\u3222")]
    [InlineData("ideographEnclosedCircle", 4, true, "\u3223")]
    [InlineData("ideographEnclosedCircle", 5, true, "\u3224")]
    [InlineData("ideographEnclosedCircle", 6, true, "\u3225")]
    [InlineData("ideographEnclosedCircle", 7, true, "\u3226")]
    [InlineData("ideographEnclosedCircle", 8, true, "\u3227")]
    [InlineData("ideographEnclosedCircle", 9, true, "\u3228")]
    [InlineData("ideographEnclosedCircle", 10, true, "\u3229")]
    [InlineData("ideographEnclosedCircle", 11, true, "11")]
    [InlineData("ideographEnclosedCircle", 12, true, "12")]
    [InlineData("ideographEnclosedCircle", 13, true, "13")]
    [InlineData("ideographEnclosedCircle", 14, true, "14")]
    [InlineData("ideographEnclosedCircle", 15, true, "15")]
    [InlineData("ideographEnclosedCircle", 16, true, "16")]
    [InlineData("ideographEnclosedCircle", 17, true, "17")]
    [InlineData("ideographEnclosedCircle", 18, true, "18")]
    [InlineData("ideographEnclosedCircle", 19, true, "19")]
    [InlineData("ideographEnclosedCircle", 20, true, "20")]
    [InlineData("ideographEnclosedCircle", 21, true, "21")]
    [InlineData("ideographEnclosedCircle", 22, true, "22")]
    [InlineData("ideographEnclosedCircle", 23, true, "23")]
    [InlineData("ideographEnclosedCircle", 24, true, "24")]
    [InlineData("ideographEnclosedCircle", 25, true, "25")]
    [InlineData("ideographEnclosedCircle", 26, true, "26")]
    [InlineData("ideographEnclosedCircle", 27, true, "27")]
    [InlineData("ideographEnclosedCircle", 28, true, "28")]
    [InlineData("ideographEnclosedCircle", 29, true, "29")]
    [InlineData("ideographEnclosedCircle", 30, true, "30")]
    [InlineData("ideographEnclosedCircle", 31, true, "31")]
    [InlineData("ideographEnclosedCircle", 32, true, "32")]
    [InlineData("ideographEnclosedCircle", 33, true, "33")]
    [InlineData("ideographEnclosedCircle", 34, true, "34")]
    [InlineData("ideographEnclosedCircle", 35, true, "35")]
    [InlineData("ideographEnclosedCircle", 36, true, "36")]
    [InlineData("ideographEnclosedCircle", 37, true, "37")]
    [InlineData("ideographEnclosedCircle", 38, true, "38")]
    [InlineData("ideographEnclosedCircle", 39, true, "39")]
    [InlineData("ideographEnclosedCircle", 40, true, "40")]
    [InlineData("ideographEnclosedCircle", 41, true, "41")]
    [InlineData("ideographEnclosedCircle", 42, true, "42")]
    [InlineData("ideographEnclosedCircle", 43, true, "43")]
    [InlineData("ideographEnclosedCircle", 44, true, "44")]
    [InlineData("ideographEnclosedCircle", 45, true, "45")]
    [InlineData("ideographEnclosedCircle", 46, true, "46")]
    [InlineData("ideographEnclosedCircle", 47, true, "47")]
    [InlineData("ideographEnclosedCircle", 48, true, "48")]
    [InlineData("ideographEnclosedCircle", 49, true, "49")]
    [InlineData("ideographEnclosedCircle", 50, true, "50")]
    [InlineData("ideographEnclosedCircle", 51, true, "51")]
    [InlineData("ideographEnclosedCircle", 52, true, "52")]
    [InlineData("ideographEnclosedCircle", 53, true, "53")]
    [InlineData("ideographEnclosedCircle", 54, true, "54")]
    [InlineData("ideographEnclosedCircle", 55, true, "55")]
    [InlineData("ideographEnclosedCircle", 56, true, "56")]
    [InlineData("ideographEnclosedCircle", 57, true, "57")]
    [InlineData("ideographEnclosedCircle", 58, true, "58")]
    [InlineData("ideographEnclosedCircle", 59, true, "59")]
    [InlineData("ideographEnclosedCircle", 60, true, "60")]
    [InlineData("ideographEnclosedCircle", 61, true, "61")]
    [InlineData("ideographEnclosedCircle", 62, true, "62")]
    [InlineData("ideographEnclosedCircle", 63, true, "63")]
    [InlineData("ideographEnclosedCircle", 64, true, "64")]
    [InlineData("ideographEnclosedCircle", 65, true, "65")]
    [InlineData("ideographEnclosedCircle", 66, true, "66")]
    [InlineData("ideographEnclosedCircle", 67, true, "67")]
    [InlineData("ideographEnclosedCircle", 68, true, "68")]
    [InlineData("ideographEnclosedCircle", 69, true, "69")]
    [InlineData("ideographEnclosedCircle", 70, true, "70")]
    [InlineData("ideographEnclosedCircle", 71, true, "71")]
    [InlineData("ideographEnclosedCircle", 72, true, "72")]
    [InlineData("ideographEnclosedCircle", 73, true, "73")]
    [InlineData("ideographEnclosedCircle", 74, true, "74")]
    [InlineData("ideographEnclosedCircle", 75, true, "75")]
    [InlineData("ideographEnclosedCircle", 76, true, "76")]
    [InlineData("ideographEnclosedCircle", 77, true, "77")]
    [InlineData("ideographEnclosedCircle", 78, true, "78")]
    [InlineData("ideographEnclosedCircle", 79, true, "79")]
    [InlineData("ideographEnclosedCircle", 80, true, "80")]
    [InlineData("ideographEnclosedCircle", 81, true, "81")]
    [InlineData("ideographEnclosedCircle", 82, true, "82")]
    [InlineData("ideographEnclosedCircle", 83, true, "83")]
    [InlineData("ideographEnclosedCircle", 84, true, "84")]
    [InlineData("ideographEnclosedCircle", 85, true, "85")]
    [InlineData("ideographEnclosedCircle", 86, true, "86")]
    [InlineData("ideographEnclosedCircle", 87, true, "87")]
    [InlineData("ideographEnclosedCircle", 88, true, "88")]
    [InlineData("ideographEnclosedCircle", 89, true, "89")]
    [InlineData("ideographEnclosedCircle", 90, true, "90")]
    [InlineData("ideographEnclosedCircle", 91, true, "91")]
    [InlineData("ideographEnclosedCircle", 92, true, "92")]
    [InlineData("ideographEnclosedCircle", 93, true, "93")]
    [InlineData("ideographEnclosedCircle", 94, true, "94")]
    [InlineData("ideographEnclosedCircle", 95, true, "95")]
    [InlineData("ideographEnclosedCircle", 96, true, "96")]
    [InlineData("ideographEnclosedCircle", 97, true, "97")]
    [InlineData("ideographEnclosedCircle", 98, true, "98")]
    [InlineData("ideographEnclosedCircle", 99, true, "99")]
    [InlineData("ideographEnclosedCircle", 100, true, "100")]
    [InlineData("ideographEnclosedCircle", 101, true, "101")]
    [InlineData("ideographEnclosedCircle", 102, true, "102")]
    [InlineData("ideographEnclosedCircle", 103, true, "103")]
    [InlineData("ideographEnclosedCircle", 104, true, "104")]
    [InlineData("ideographEnclosedCircle", 105, true, "105")]
    [InlineData("ideographEnclosedCircle", 106, true, "106")]
    [InlineData("ideographEnclosedCircle", 107, true, "107")]
    [InlineData("ideographEnclosedCircle", 108, true, "108")]
    [InlineData("ideographEnclosedCircle", 109, true, "109")]
    [InlineData("ideographEnclosedCircle", 110, true, "110")]
    [InlineData("ideographEnclosedCircle", 111, true, "111")]
    [InlineData("ideographEnclosedCircle", 112, true, "112")]
    [InlineData("ideographEnclosedCircle", 113, true, "113")]
    [InlineData("ideographEnclosedCircle", 114, true, "114")]
    [InlineData("ideographEnclosedCircle", 115, true, "115")]
    [InlineData("ideographEnclosedCircle", 116, true, "116")]
    [InlineData("ideographEnclosedCircle", 117, true, "117")]
    [InlineData("ideographEnclosedCircle", 118, true, "118")]
    [InlineData("ideographEnclosedCircle", 119, true, "119")]
    [InlineData("ideographEnclosedCircle", 120, true, "120")]
    [InlineData("ideographEnclosedCircle", 121, true, "121")]
    [InlineData("ideographEnclosedCircle", 420, true, "420")]
    [InlineData("ideographEnclosedCircle", 421, true, "421")]
    [InlineData("ideographEnclosedCircle", 32767, true, "32767")]
    [InlineData("ideographEnclosedCircle", 32768, true, "32768")]
    [InlineData("ideographEnclosedCircle", 65535, true, "65535")]
    [InlineData("ideographEnclosedCircle", 999999, true, "999999")]
    [InlineData("ideographEnclosedCircle", 1000000, true, "1000000")]
    [InlineData("ideographEnclosedCircle", 2147483647, true, "2147483647")]
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
