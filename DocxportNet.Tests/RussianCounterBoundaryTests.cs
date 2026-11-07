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

public sealed class RussianCounterBoundaryTests
{
    [Theory]
    [InlineData("russianLower")]
    [InlineData("russianUpper")]
    public void SharedFieldFormatterRetainsItsRepetitionLimit(string format)
    {
        Assert.Equal(new string(format == "russianLower" ? '\u0430' : '\u0410', 31), DxpListMarkerFormatter.FormatNumber(871, new NumberFormatValues(format)));
        Assert.Equal("7425", DxpListMarkerFormatter.FormatNumber(7425, new NumberFormatValues(format)));
    }

    // Captured from visible Word PDFs, including negative abstract starts vs start overrides.
    [Theory]
    [InlineData("russianLower", 1, false, "\u0430")]
    [InlineData("russianLower", 29, false, "\u044f")]
    [InlineData("russianLower", 30, false, "\u0430\u0430")]
    [InlineData("russianLower", 869, false, "\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e\u044e")]
    [InlineData("russianLower", 870, false, "\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f")]
    [InlineData("russianLower", 871, false, "\u0430")]
    [InlineData("russianLower", 899, false, "\u044f")]
    [InlineData("russianLower", 7424, false, "\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f\u044f")]
    [InlineData("russianLower", 7425, false, "\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430\u0430")]
    [InlineData("russianLower", 2147483647, false, "\u0436")]
    [InlineData("russianUpper", 1, false, "\u0410")]
    [InlineData("russianUpper", 29, false, "\u042f")]
    [InlineData("russianUpper", 30, false, "\u0410\u0410")]
    [InlineData("russianUpper", 869, false, "\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e\u042e")]
    [InlineData("russianUpper", 870, false, "\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f")]
    [InlineData("russianUpper", 871, false, "\u0410")]
    [InlineData("russianUpper", 899, false, "\u042f")]
    [InlineData("russianUpper", 7424, false, "\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f\u042f")]
    [InlineData("russianUpper", 7425, false, "\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410\u0410")]
    [InlineData("russianUpper", 2147483647, false, "\u0416")]
    [InlineData("russianLower", -2147483648, false, null)]
    [InlineData("russianLower", -1, false, null)]
    [InlineData("russianUpper", -2147483648, false, null)]
    [InlineData("russianUpper", -1, false, null)]
    [InlineData("russianLower", -2147483648, true, "")]
    [InlineData("russianLower", -1, true, "")]
    [InlineData("russianUpper", -2147483648, true, "")]
    [InlineData("russianUpper", -1, true, "")]
    public void RepeatedMarkersMatchWordAtBoundaries(string format, int start, bool startOverride, string? expected)
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
    public static IEnumerable<object[]> WordSequences()
    {
        yield return new object[] { "russianLower", 0, false, new string[] { " ", "\u0430", "\u0431", "\u0432", "\u0433", "\u0434" } };
        yield return new object[] { "russianUpper", 0, false, new string[] { " ", "\u0410", "\u0411", "\u0412", "\u0413", "\u0414" } };
    }

    [Theory]
    [MemberData(nameof(WordSequences))]
    public void SuccessiveMarkersMatchWordAcrossIntegerBoundaries(string format, int start, bool startOverride, string[] expected)
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
