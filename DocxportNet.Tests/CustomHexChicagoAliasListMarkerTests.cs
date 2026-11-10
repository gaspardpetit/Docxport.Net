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

public sealed class CustomHexChicagoAliasListMarkerTests
{
    // Exact hexadecimal and Chicago aliases, checked against retained Word captures.
    [Theory]
    [InlineData("chicago", -2147483648, 0, "")]
    [InlineData("chicago", -1, 0, "")]
    [InlineData("chicago", 0, 0, "")]
    [InlineData("chicago", 1, 0, "*")]
    [InlineData("chicago", 9, 0, "***")]
    [InlineData("chicago", 10, 0, "\u2020\u2020\u2020")]
    [InlineData("chicago", 99, 0, "\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021")]
    [InlineData("chicago", 100, 0, "\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7")]
    [InlineData("chicago", 999, 0, "\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021")]
    [InlineData("chicago", 1000, 0, "\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7")]
    [InlineData("chicago", 9999, 0, "\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021")]
    [InlineData("chicago", 10000, 0, "\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7\u00a7")]
    [InlineData("chicago", 32767, 0, "\u2021\u2021")]
    [InlineData("chicago", 32768, 0, "\u00a7\u00a7")]
    [InlineData("chicago", 65535, 0, "\u2021\u2021\u2021\u2021")]
    [InlineData("chicago", 999999, 0, "\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021\u2021")]
    [InlineData("chicago", 2147483647, 0, "\u2021\u2021")]
    [InlineData("hex", -2147483648, 0, "")]
    [InlineData("hex", -1, 0, "")]
    [InlineData("hex", 0, 0, "0")]
    [InlineData("hex", 1, 0, "1")]
    [InlineData("hex", 9, 0, "9")]
    [InlineData("hex", 10, 0, "A")]
    [InlineData("hex", 99, 0, "63")]
    [InlineData("hex", 100, 0, "64")]
    [InlineData("hex", 999, 0, "3E7")]
    [InlineData("hex", 1000, 0, "3E8")]
    [InlineData("hex", 9999, 0, "270F")]
    [InlineData("hex", 10000, 0, "2710")]
    [InlineData("hex", 32767, 0, "7FFF")]
    [InlineData("hex", 32768, 0, "8000")]
    [InlineData("hex", 65535, 0, "FFFF")]
    [InlineData("hex", 999999, 0, "")]
    [InlineData("hex", 2147483647, 0, "")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var abstractXml = name switch
        {
            "chicago" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"*, \u2020, \u2021, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "hex" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"1, A, B, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
