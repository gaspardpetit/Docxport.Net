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

public sealed class KanaListMarkerTests
{
    [Theory]
    [InlineData("aiueo")]
    [InlineData("aiueoFullWidth")]
    [InlineData("iroha")]
    [InlineData("irohaFullWidth")]
    public void SharedFieldFormatterRetainsDecimalFallback(string format)
    {
        Assert.Equal("1", DxpListMarkerFormatter.FormatNumber(1, new NumberFormatValues(format)));
        Assert.Equal("49", DxpListMarkerFormatter.FormatNumber(49, new NumberFormatValues(format)));
    }

    // Word ListString sequences, corroborated by Word PDF export.
    public static IEnumerable<object[]> WordSequences()
    {
        yield return new object[] { "aiueo", 0, false, new string[] { "0", "\uff71", "\uff72", "\uff73", "\uff74", "\uff75", "\uff76", "\uff77", "\uff78", "\uff79", "\uff7a", "\uff7b", "\uff7c", "\uff7d", "\uff7e", "\uff7f", "\uff80", "\uff81", "\uff82", "\uff83", "\uff84", "\uff85", "\uff86", "\uff87", "\uff88", "\uff89", "\uff8a", "\uff8b", "\uff8c", "\uff8d", "\uff8e", "\uff8f", "\uff90", "\uff91", "\uff92", "\uff93", "\uff94", "\uff95", "\uff96", "\uff97", "\uff98", "\uff99", "\uff9a", "\uff9b", "\uff9c", "\uff66", "\uff9d", "\uff71", "\uff72", "\uff73", "\uff74", "\uff75", "\uff76", "\uff77", "\uff78", "\uff79", "\uff7a", "\uff7b", "\uff7c", "\uff7d", "\uff7e" } };
        yield return new object[] { "aiueoFullWidth", 0, false, new string[] { "0", "\u30a2", "\u30a4", "\u30a6", "\u30a8", "\u30aa", "\u30ab", "\u30ad", "\u30af", "\u30b1", "\u30b3", "\u30b5", "\u30b7", "\u30b9", "\u30bb", "\u30bd", "\u30bf", "\u30c1", "\u30c4", "\u30c6", "\u30c8", "\u30ca", "\u30cb", "\u30cc", "\u30cd", "\u30ce", "\u30cf", "\u30d2", "\u30d5", "\u30d8", "\u30db", "\u30de", "\u30df", "\u30e0", "\u30e1", "\u30e2", "\u30e4", "\u30e6", "\u30e8", "\u30e9", "\u30ea", "\u30eb", "\u30ec", "\u30ed", "\u30ef", "\u30f2", "\u30f3", "\u30a2", "\u30a4", "\u30a6", "\u30a8", "\u30aa", "\u30ab", "\u30ad", "\u30af", "\u30b1", "\u30b3", "\u30b5", "\u30b7", "\u30b9", "\u30bb" } };
        yield return new object[] { "iroha", 0, false, new string[] { "0", "\uff72", "\uff9b", "\uff8a", "\uff86", "\uff8e", "\uff8d", "\uff84", "\uff81", "\uff98", "\uff87", "\uff99", "\uff66", "\uff9c", "\uff76", "\uff96", "\uff80", "\uff9a", "\uff7f", "\uff82", "\uff88", "\uff85", "\uff97", "\uff91", "\uff73", "\u30f0", "\uff89", "\uff75", "\uff78", "\uff94", "\uff8f", "\uff79", "\uff8c", "\uff7a", "\uff74", "\uff83", "\uff71", "\uff7b", "\uff77", "\uff95", "\uff92", "\uff90", "\uff7c", "\u30f1", "\uff8b", "\uff93", "\uff7e", "\uff7d", "\uff9d", "\uff72", "\uff9b", "\uff8a", "\uff86", "\uff8e", "\uff8d", "\uff84", "\uff81", "\uff98", "\uff87", "\uff99", "\uff66" } };
        yield return new object[] { "irohaFullWidth", 0, false, new string[] { "0", "\u30a4", "\u30ed", "\u30cf", "\u30cb", "\u30db", "\u30d8", "\u30c8", "\u30c1", "\u30ea", "\u30cc", "\u30eb", "\u30f2", "\u30ef", "\u30ab", "\u30e8", "\u30bf", "\u30ec", "\u30bd", "\u30c4", "\u30cd", "\u30ca", "\u30e9", "\u30e0", "\u30a6", "\u30f0", "\u30ce", "\u30aa", "\u30af", "\u30e4", "\u30de", "\u30b1", "\u30d5", "\u30b3", "\u30a8", "\u30c6", "\u30a2", "\u30b5", "\u30ad", "\u30e6", "\u30e1", "\u30df", "\u30b7", "\u30f1", "\u30d2", "\u30e2", "\u30bb", "\u30b9", "\u30f3", "\u30a4", "\u30ed", "\u30cf", "\u30cb", "\u30db", "\u30d8", "\u30c8", "\u30c1", "\u30ea", "\u30cc", "\u30eb", "\u30f2" } };
    }

    [Theory]
    [MemberData(nameof(WordSequences))]
    public void SuccessiveMarkersMatchWordAcrossKanaCycles(string format, int start, bool startOverride, string[] expected)
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
