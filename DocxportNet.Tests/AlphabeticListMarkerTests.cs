using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Walker;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class AlphabeticListMarkerTests
{
    // Golden labels captured from Word, including the 30-letter wrap boundary.
    public static IEnumerable<object[]> WordCases()
    {
        yield return new object[] { 1000, "lowerLetter", "lllllllll" };
        yield return new object[] { 100000, "lowerLetter", "ddddddd" };
        yield return new object[] { 10345021, "lowerLetter", "kkkkkkkkkkkkkkkkkkkkkkkkkk" };
        yield return new object[] { 26, "lowerLetter", "z" };
        yield return new object[] { 6656, "lowerLetter", "zzzzzzzzzzzzzzzz" };
        yield return new object[] { 6657, "lowerLetter", "aaaaaaaaaaaaaaaaa" };
        yield return new object[] { 676, "lowerLetter", "zzzzzzzzzzzzzzzzzzzzzzzzzz" };
        yield return new object[] { 677, "lowerLetter", "aaaaaaaaaaaaaaaaaaaaaaaaaaa" };
        yield return new object[] { 779, "lowerLetter", "yyyyyyyyyyyyyyyyyyyyyyyyyyyyyy" };
        yield return new object[] { 780, "lowerLetter", "zzzzzzzzzzzzzzzzzzzzzzzzzzzzzz" };
        yield return new object[] { 781, "lowerLetter", "a" };
        yield return new object[] { 782, "lowerLetter", "b" };
        yield return new object[] { 1000, "upperLetter", "LLLLLLLLL" };
        yield return new object[] { 780, "upperLetter", "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ" };
        yield return new object[] { 781, "upperLetter", "A" };
        yield return new object[] { 26, "upperLetter", "Z" };
        yield return new object[] { 27, "upperLetter", "AA" };
        yield return new object[] { 27, "lowerLetter", "aa" };
        yield return new object[] { 28, "upperLetter", "BB" };
        yield return new object[] { 28, "lowerLetter", "bb" };
        yield return new object[] { 51, "upperLetter", "YY" };
        yield return new object[] { 51, "lowerLetter", "yy" };
        yield return new object[] { 52, "upperLetter", "ZZ" };
        yield return new object[] { 52, "lowerLetter", "zz" };
        yield return new object[] { 53, "upperLetter", "AAA" };
        yield return new object[] { 53, "lowerLetter", "aaa" };
        yield return new object[] { 54, "upperLetter", "BBB" };
        yield return new object[] { 54, "lowerLetter", "bbb" };
        yield return new object[] { 701, "upperLetter", "YYYYYYYYYYYYYYYYYYYYYYYYYYY" };
        yield return new object[] { 701, "lowerLetter", "yyyyyyyyyyyyyyyyyyyyyyyyyyy" };
        yield return new object[] { 702, "upperLetter", "ZZZZZZZZZZZZZZZZZZZZZZZZZZZ" };
        yield return new object[] { 702, "lowerLetter", "zzzzzzzzzzzzzzzzzzzzzzzzzzz" };
    }

    [Theory]
    [MemberData(nameof(WordCases))]
    public void TextExportersMatchWordAlphabeticLists(int start, string format, string expected)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
            new AbstractNum(new Level(new StartNumberingValue { Val = start },
                new NumberingFormat { Val = new NumberFormatValues(format) },
                new LevelText { Val = "[%1]" }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
        main.Document = new Document(new Body(new Paragraph(new ParagraphProperties(
            new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
            new Run(new Text("Item")))));
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document));
        var before = main.NumberingDefinitionsPart.Numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(document, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            Assert.Equal("[" + expected + "] Item", output.Trim());
        }
        Assert.Equal(before, main.NumberingDefinitionsPart.Numbering.OuterXml);
    }

    [Theory]
    [InlineData(28, "AB", "ab")]
    [InlineData(53, "BA", "ba")]
    public void GenericFieldFormatterRetainsPositionalLetters(int value, string upper, string lower)
    {
        Assert.Equal(upper, DxpListMarkerFormatter.FormatNumber(value, NumberFormatValues.UpperLetter));
        Assert.Equal(lower, DxpListMarkerFormatter.FormatNumber(value, NumberFormatValues.LowerLetter));
    }
}
