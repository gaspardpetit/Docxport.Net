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

public sealed class CustomKoreanAlphabetAliasListMarkerTests
{
    // Word captures for the Ganada and Chosung custom aliases.
    [Theory]
    [InlineData("chosung", -2147483648, 0, "")]
    [InlineData("chosung", -1, 0, "-1")]
    [InlineData("chosung", 0, 0, "0")]
    [InlineData("chosung", 1, 0, "\u3131")]
    [InlineData("chosung", 9, 0, "\u3148")]
    [InlineData("chosung", 10, 0, "\u314a")]
    [InlineData("chosung", 99, 0, "\u3131")]
    [InlineData("chosung", 100, 0, "\u3134")]
    [InlineData("chosung", 999, 0, "\u3141")]
    [InlineData("chosung", 1000, 0, "\u3142")]
    [InlineData("chosung", 9999, 0, "\u3137")]
    [InlineData("chosung", 10000, 0, "\u3139")]
    [InlineData("chosung", 32767, 0, "\u3145")]
    [InlineData("chosung", 32768, 0, "\u3147")]
    [InlineData("chosung", 65535, 0, "\u3131")]
    [InlineData("chosung", 999999, 0, "\u3145")]
    [InlineData("chosung", 2147483647, 0, "\u3131")]
    [InlineData("ganada", -2147483648, 0, "")]
    [InlineData("ganada", -1, 0, "-1")]
    [InlineData("ganada", 0, 0, "0")]
    [InlineData("ganada", 1, 0, "\uac00")]
    [InlineData("ganada", 9, 0, "\uc790")]
    [InlineData("ganada", 10, 0, "\ucc28")]
    [InlineData("ganada", 99, 0, "\uac00")]
    [InlineData("ganada", 100, 0, "\ub098")]
    [InlineData("ganada", 999, 0, "\ub9c8")]
    [InlineData("ganada", 1000, 0, "\ubc14")]
    [InlineData("ganada", 9999, 0, "\ub2e4")]
    [InlineData("ganada", 10000, 0, "\ub77c")]
    [InlineData("ganada", 32767, 0, "\uc0ac")]
    [InlineData("ganada", 32768, 0, "\uc544")]
    [InlineData("ganada", 65535, 0, "\uac00")]
    [InlineData("ganada", 999999, 0, "\uc0ac")]
    [InlineData("ganada", 2147483647, 0, "\uac00")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var abstractXml = name switch
        {
            "chosung" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u3131, \u3134, \u3137, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ganada" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\uac00, \ub098, \ub2e4, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(new AbstractNum(abstractXml),
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = level }) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        var item = new Paragraph(new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = level }, new NumberingId { Val = 2 })), new Run(new Text("Body")));
        main.Document = new Document(new Body(item));
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
