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

public sealed class EnclosedCounterBoundaryTests
{
    [Fact]
    public void SharedCircledFormatterRetainsAdditionalFieldSymbols()
    {
        Assert.Equal("\u3251", DxpListMarkerFormatter.FormatNumber(21, NumberFormatValues.DecimalEnclosedCircle));
        Assert.Equal("\u32b1", DxpListMarkerFormatter.FormatNumber(36, NumberFormatValues.DecimalEnclosedCircle));
    }

    // Word ListString sequences, corroborated by Word PDF export.
    public static IEnumerable<object[]> WordSequences()
    {
        yield return new object[] { "decimalEnclosedCircle", 0, false, new string[] { "0", "\u2460", "\u2461", "\u2462", "\u2463", "\u2464", "\u2465", "\u2466", "\u2467", "\u2468", "\u2469", "\u246a", "\u246b", "\u246c", "\u246d", "\u246e", "\u246f", "\u2470", "\u2471", "\u2472", "\u2473", "21", "22" } };
        yield return new object[] { "decimalEnclosedCircle", 34, false, new string[] { "34", "35", "36", "37", "38", "39", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "50", "51", "52" } };
        yield return new object[] { "decimalEnclosedCircleChinese", 0, false, new string[] { "0", "\u2460", "\u2461", "\u2462", "\u2463", "\u2464", "\u2465", "\u2466", "\u2467", "\u2468", "\u2469", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22" } };
        yield return new object[] { "decimalEnclosedFullstop", 0, false, new string[] { "0", "\u2488", "\u2489", "\u248a", "\u248b", "\u248c", "\u248d", "\u248e", "\u248f", "\u2490", "\u2491", "\u2492", "\u2493", "\u2494", "\u2495", "\u2496", "\u2497", "\u2498", "\u2499", "\u249a", "\u249b", "21", "22" } };
        yield return new object[] { "decimalEnclosedParen", 0, false, new string[] { "0", "\u2474", "\u2475", "\u2476", "\u2477", "\u2478", "\u2479", "\u247a", "\u247b", "\u247c", "\u247d", "\u247e", "\u247f", "\u2480", "\u2481", "\u2482", "\u2483", "\u2484", "\u2485", "\u2486", "\u2487", "21", "22" } };
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
