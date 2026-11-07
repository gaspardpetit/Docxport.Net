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

public sealed class RomanCounterBoundaryTests
{
    [Fact]
    public void SharedRomanFormatterRetainsFieldBehavior()
    {
        Assert.Equal(new string('M', 32) + "DCCLXVIII",
            DxpListMarkerFormatter.FormatNumber(32768, NumberFormatValues.UpperRoman));
        Assert.Equal("0", DxpListMarkerFormatter.FormatNumber(0, NumberFormatValues.LowerRoman));
        Assert.Equal("-1", DxpListMarkerFormatter.FormatNumber(-1, NumberFormatValues.UpperRoman));
    }

    // Word ListString sequences, corroborated by Word PDF export.
    public static IEnumerable<object[]> WordSequences()
    {
        yield return new object[] { "upperRoman", 2147483646, false, new string[] { "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVII", "I", "", "", "", "" } };
        yield return new object[] { "upperRoman", 2147483646, true, new string[] { "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVII", "I", "", "", "", "" } };
        yield return new object[] { "upperRoman", 32766, false, new string[] { "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVI", "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVII", "I", "II" } };
        yield return new object[] { "upperRoman", 65533, false, new string[] { "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVI", "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVII", "I", "II" } };

        yield return new object[] { "upperRoman", 0, false, new string[] { " ", "I", "II", "III", "IV", "V" } };
        yield return new object[] { "upperRoman", -2147483648, true, new string[] { "", "", "", "", "", "" } };
        yield return new object[] { "upperRoman", -1, true, new string[] { "", " ", "I", "II", "III", "IV" } };
        yield return new object[] { "upperRoman", 0, true, new string[] { " ", "I", "II", "III", "IV", "V" } };
        yield return new object[] { "lowerRoman", 0, false, new string[] { " ", "i", "ii", "iii", "iv", "v" } };
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
