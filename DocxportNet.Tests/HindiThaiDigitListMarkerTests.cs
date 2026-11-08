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

public sealed class HindiThaiDigitListMarkerTests
{
    [Theory]
    [InlineData("hindiNumbers")]
    [InlineData("thaiNumbers")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("32768", DxpListMarkerFormatter.FormatNumber(32768, new NumberFormatValues(format)));
    }

    // Word ListString captures with PDF artifacts retained; negative values are start overrides.
    [Theory]
    [InlineData("hindiNumbers", -2147483648, true, "")]
    [InlineData("hindiNumbers", -1, true, "")]
    [InlineData("hindiNumbers", 0, true, "\u0966")]
    [InlineData("hindiNumbers", 1, true, "\u0967")]
    [InlineData("hindiNumbers", 2, true, "\u0968")]
    [InlineData("hindiNumbers", 9, true, "\u096f")]
    [InlineData("hindiNumbers", 10, true, "\u0967\u0966")]
    [InlineData("hindiNumbers", 21, true, "\u0968\u0967")]
    [InlineData("hindiNumbers", 100, true, "\u0967\u0966\u0966")]
    [InlineData("hindiNumbers", 32766, true, "\u0969\u0968\u096d\u096c\u096c")]
    [InlineData("hindiNumbers", 32767, true, "\u0969\u0968\u096d\u096c\u096d")]
    [InlineData("hindiNumbers", 32768, true, "\u0967")]
    [InlineData("hindiNumbers", 65534, true, "\u0969\u0968\u096d\u096c\u096d")]
    [InlineData("hindiNumbers", 65535, true, "\u0967")]
    [InlineData("hindiNumbers", 999999, true, "\u0967\u096c\u096f\u096e\u096f")]
    [InlineData("hindiNumbers", 1000000, true, "\u0967\u096c\u096f\u096f\u0966")]
    [InlineData("hindiNumbers", 2147483647, true, "\u0967")]
    [InlineData("thaiNumbers", -2147483648, true, "")]
    [InlineData("thaiNumbers", -1, true, "")]
    [InlineData("thaiNumbers", 0, true, "\u0e50")]
    [InlineData("thaiNumbers", 1, true, "\u0e51")]
    [InlineData("thaiNumbers", 2, true, "\u0e52")]
    [InlineData("thaiNumbers", 9, true, "\u0e59")]
    [InlineData("thaiNumbers", 10, true, "\u0e51\u0e50")]
    [InlineData("thaiNumbers", 21, true, "\u0e52\u0e51")]
    [InlineData("thaiNumbers", 100, true, "\u0e51\u0e50\u0e50")]
    [InlineData("thaiNumbers", 32766, true, "\u0e53\u0e52\u0e57\u0e56\u0e56")]
    [InlineData("thaiNumbers", 32767, true, "\u0e53\u0e52\u0e57\u0e56\u0e57")]
    [InlineData("thaiNumbers", 32768, true, "\u0e51")]
    [InlineData("thaiNumbers", 65534, true, "\u0e53\u0e52\u0e57\u0e56\u0e57")]
    [InlineData("thaiNumbers", 65535, true, "\u0e51")]
    [InlineData("thaiNumbers", 999999, true, "\u0e51\u0e56\u0e59\u0e58\u0e59")]
    [InlineData("thaiNumbers", 1000000, true, "\u0e51\u0e56\u0e59\u0e59\u0e50")]
    [InlineData("thaiNumbers", 2147483647, true, "\u0e51")]
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
