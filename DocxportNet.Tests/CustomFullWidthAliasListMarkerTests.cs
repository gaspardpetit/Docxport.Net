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

public sealed class CustomFullWidthAliasListMarkerTests
{
    // Full-width alias and direct formats, checked against Word captures.
    [Theory]
    [InlineData("decimalFullWidth", -2147483648, 0, "")]
    [InlineData("decimalFullWidth", -1, 0, "-\uff11")]
    [InlineData("decimalFullWidth", 0, 0, "\uff10")]
    [InlineData("decimalFullWidth", 1, 0, "\uff11")]
    [InlineData("decimalFullWidth", 9, 0, "\uff19")]
    [InlineData("decimalFullWidth", 10, 0, "\uff11\uff10")]
    [InlineData("decimalFullWidth", 99, 0, "\uff19\uff19")]
    [InlineData("decimalFullWidth", 100, 0, "\uff11\uff10\uff10")]
    [InlineData("decimalFullWidth", 999, 0, "\uff19\uff19\uff19")]
    [InlineData("decimalFullWidth", 1000, 0, "\uff11\uff10\uff10\uff10")]
    [InlineData("decimalFullWidth", 9999, 0, "\uff19\uff19\uff19\uff19")]
    [InlineData("decimalFullWidth", 10000, 0, "\uff11\uff10\uff10\uff10\uff10")]
    [InlineData("decimalFullWidth", 32767, 0, "\uff13\uff12\uff17\uff16\uff17")]
    [InlineData("decimalFullWidth", 32768, 0, "\uff13\uff12\uff17\uff16\uff18")]
    [InlineData("decimalFullWidth", 65535, 0, "\uff16\uff15\uff15\uff13\uff15")]
    [InlineData("decimalFullWidth", 999999, 0, "\uff19\uff19\uff19\uff19\uff19\uff19")]
    [InlineData("decimalFullWidth", 2147483647, 0, "\uff12\uff11\uff14\uff17\uff14\uff18\uff13\uff16\uff14\uff17")]
    [InlineData("direct-decimalFullWidth", -2147483648, 0, "")]
    [InlineData("direct-decimalFullWidth", -1, 0, "-\uff11")]
    [InlineData("direct-decimalFullWidth", 0, 0, "\uff10")]
    [InlineData("direct-decimalFullWidth", 1, 0, "\uff11")]
    [InlineData("direct-decimalFullWidth", 9, 0, "\uff19")]
    [InlineData("direct-decimalFullWidth", 10, 0, "\uff11\uff10")]
    [InlineData("direct-decimalFullWidth", 99, 0, "\uff19\uff19")]
    [InlineData("direct-decimalFullWidth", 100, 0, "\uff11\uff10\uff10")]
    [InlineData("direct-decimalFullWidth", 999, 0, "\uff19\uff19\uff19")]
    [InlineData("direct-decimalFullWidth", 1000, 0, "\uff11\uff10\uff10\uff10")]
    [InlineData("direct-decimalFullWidth", 9999, 0, "\uff19\uff19\uff19\uff19")]
    [InlineData("direct-decimalFullWidth", 10000, 0, "\uff11\uff10\uff10\uff10\uff10")]
    [InlineData("direct-decimalFullWidth", 32767, 0, "\uff13\uff12\uff17\uff16\uff17")]
    [InlineData("direct-decimalFullWidth", 32768, 0, "\uff13\uff12\uff17\uff16\uff18")]
    [InlineData("direct-decimalFullWidth", 65535, 0, "\uff16\uff15\uff15\uff13\uff15")]
    [InlineData("direct-decimalFullWidth", 999999, 0, "\uff19\uff19\uff19\uff19\uff19\uff19")]
    [InlineData("direct-decimalFullWidth", 2147483647, 0, "\uff12\uff11\uff14\uff17\uff14\uff18\uff13\uff16\uff14\uff17")]
    [InlineData("direct-decimalFullWidth2", -2147483648, 0, "")]
    [InlineData("direct-decimalFullWidth2", -1, 0, "")]
    [InlineData("direct-decimalFullWidth2", 0, 0, "\uff10")]
    [InlineData("direct-decimalFullWidth2", 1, 0, "\uff11")]
    [InlineData("direct-decimalFullWidth2", 9, 0, "\uff19")]
    [InlineData("direct-decimalFullWidth2", 10, 0, "\uff11\uff10")]
    [InlineData("direct-decimalFullWidth2", 99, 0, "\uff19\uff19")]
    [InlineData("direct-decimalFullWidth2", 100, 0, "\uff11\uff10\uff10")]
    [InlineData("direct-decimalFullWidth2", 999, 0, "\uff19\uff19\uff19")]
    [InlineData("direct-decimalFullWidth2", 1000, 0, "\uff11\uff10\uff10\uff10")]
    [InlineData("direct-decimalFullWidth2", 9999, 0, "\uff19\uff19\uff19\uff19")]
    [InlineData("direct-decimalFullWidth2", 10000, 0, "\uff11\uff10\uff10\uff10\uff10")]
    [InlineData("direct-decimalFullWidth2", 32767, 0, "\uff13\uff12\uff17\uff16\uff17")]
    [InlineData("direct-decimalFullWidth2", 32768, 0, "\uff13\uff12\uff17\uff16\uff18")]
    [InlineData("direct-decimalFullWidth2", 65535, 0, "\uff16\uff15\uff15\uff13\uff15")]
    [InlineData("direct-decimalFullWidth2", 999999, 0, "\uff19\uff19\uff19\uff19\uff19\uff19")]
    [InlineData("direct-decimalFullWidth2", 2147483647, 0, "\uff12\uff11\uff14\uff17\uff14\uff18\uff13\uff16\uff14\uff17")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat == "decimalFullWidth2" ? "decimalFullWidth" : directFormat ?? name) switch
        {
            "decimalFullWidth" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\uff11, \uff12, \uff13, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
