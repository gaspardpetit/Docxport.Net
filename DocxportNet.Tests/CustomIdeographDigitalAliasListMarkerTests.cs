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

public sealed class CustomIdeographDigitalAliasListMarkerTests
{
    // Ideographic digit alias and direct formats, checked against Word captures.
    [Theory]
    [InlineData("ideographDigital", -2147483648, 0, "")]
    [InlineData("ideographDigital", -1, 0, "")]
    [InlineData("ideographDigital", 0, 0, "\u3007")]
    [InlineData("ideographDigital", 1, 0, "\u4e00")]
    [InlineData("ideographDigital", 9, 0, "\u4e5d")]
    [InlineData("ideographDigital", 10, 0, "\u4e00\u3007")]
    [InlineData("ideographDigital", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("ideographDigital", 100, 0, "\u4e00\u3007\u3007")]
    [InlineData("ideographDigital", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("ideographDigital", 1000, 0, "\u4e00\u3007\u3007\u3007")]
    [InlineData("ideographDigital", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("ideographDigital", 10000, 0, "\u4e00\u3007\u3007\u3007\u3007")]
    [InlineData("ideographDigital", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("ideographDigital", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("ideographDigital", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("ideographDigital", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("ideographDigital", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    [InlineData("direct-ideographDigital", -2147483648, 0, "")]
    [InlineData("direct-ideographDigital", -1, 0, "")]
    [InlineData("direct-ideographDigital", 0, 0, "\u3007")]
    [InlineData("direct-ideographDigital", 1, 0, "\u4e00")]
    [InlineData("direct-ideographDigital", 9, 0, "\u4e5d")]
    [InlineData("direct-ideographDigital", 10, 0, "\u4e00\u3007")]
    [InlineData("direct-ideographDigital", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("direct-ideographDigital", 100, 0, "\u4e00\u3007\u3007")]
    [InlineData("direct-ideographDigital", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-ideographDigital", 1000, 0, "\u4e00\u3007\u3007\u3007")]
    [InlineData("direct-ideographDigital", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-ideographDigital", 10000, 0, "\u4e00\u3007\u3007\u3007\u3007")]
    [InlineData("direct-ideographDigital", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("direct-ideographDigital", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("direct-ideographDigital", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("direct-ideographDigital", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-ideographDigital", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "ideographDigital" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u4e00\u3007, \u4e00\u3007\u3007, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abstractNum = new AbstractNum(abstractXml);
        if (directFormat != null) abstractNum.GetFirstChild<Level>()!.NumberingFormat = new NumberingFormat { Val = new NumberFormatValues(directFormat) };
        var numbering = new Numbering(abstractNum,
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
