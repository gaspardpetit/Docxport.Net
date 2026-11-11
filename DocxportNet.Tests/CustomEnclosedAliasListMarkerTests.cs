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

public sealed class CustomEnclosedAliasListMarkerTests
{
    // Exact enclosed-number aliases, checked against Word captures.
    [Theory]
    [InlineData("decimalEnclosedCircle", -2147483648, 0, "")]
    [InlineData("decimalEnclosedCircle", -1, 0, "-1")]
    [InlineData("decimalEnclosedCircle", 0, 0, "0")]
    [InlineData("decimalEnclosedCircle", 1, 0, "\u2460")]
    [InlineData("decimalEnclosedCircle", 9, 0, "\u2468")]
    [InlineData("decimalEnclosedCircle", 10, 0, "\u2469")]
    [InlineData("decimalEnclosedCircle", 99, 0, "99")]
    [InlineData("decimalEnclosedCircle", 100, 0, "100")]
    [InlineData("decimalEnclosedCircle", 999, 0, "999")]
    [InlineData("decimalEnclosedCircle", 1000, 0, "1000")]
    [InlineData("decimalEnclosedCircle", 9999, 0, "9999")]
    [InlineData("decimalEnclosedCircle", 10000, 0, "10000")]
    [InlineData("decimalEnclosedCircle", 32767, 0, "32767")]
    [InlineData("decimalEnclosedCircle", 32768, 0, "32768")]
    [InlineData("decimalEnclosedCircle", 65535, 0, "65535")]
    [InlineData("decimalEnclosedCircle", 999999, 0, "999999")]
    [InlineData("decimalEnclosedCircle", 2147483647, 0, "2147483647")]
    [InlineData("decimalEnclosedFullstop", -2147483648, 0, "")]
    [InlineData("decimalEnclosedFullstop", -1, 0, "-1")]
    [InlineData("decimalEnclosedFullstop", 0, 0, "0")]
    [InlineData("decimalEnclosedFullstop", 1, 0, "\u2488")]
    [InlineData("decimalEnclosedFullstop", 9, 0, "\u2490")]
    [InlineData("decimalEnclosedFullstop", 10, 0, "\u2491")]
    [InlineData("decimalEnclosedFullstop", 99, 0, "99")]
    [InlineData("decimalEnclosedFullstop", 100, 0, "100")]
    [InlineData("decimalEnclosedFullstop", 999, 0, "999")]
    [InlineData("decimalEnclosedFullstop", 1000, 0, "1000")]
    [InlineData("decimalEnclosedFullstop", 9999, 0, "9999")]
    [InlineData("decimalEnclosedFullstop", 10000, 0, "10000")]
    [InlineData("decimalEnclosedFullstop", 32767, 0, "32767")]
    [InlineData("decimalEnclosedFullstop", 32768, 0, "32768")]
    [InlineData("decimalEnclosedFullstop", 65535, 0, "65535")]
    [InlineData("decimalEnclosedFullstop", 999999, 0, "999999")]
    [InlineData("decimalEnclosedFullstop", 2147483647, 0, "2147483647")]
    [InlineData("decimalEnclosedParen", -2147483648, 0, "")]
    [InlineData("decimalEnclosedParen", -1, 0, "-1")]
    [InlineData("decimalEnclosedParen", 0, 0, "0")]
    [InlineData("decimalEnclosedParen", 1, 0, "\u2474")]
    [InlineData("decimalEnclosedParen", 9, 0, "\u247c")]
    [InlineData("decimalEnclosedParen", 10, 0, "\u247d")]
    [InlineData("decimalEnclosedParen", 99, 0, "99")]
    [InlineData("decimalEnclosedParen", 100, 0, "100")]
    [InlineData("decimalEnclosedParen", 999, 0, "999")]
    [InlineData("decimalEnclosedParen", 1000, 0, "1000")]
    [InlineData("decimalEnclosedParen", 9999, 0, "9999")]
    [InlineData("decimalEnclosedParen", 10000, 0, "10000")]
    [InlineData("decimalEnclosedParen", 32767, 0, "32767")]
    [InlineData("decimalEnclosedParen", 32768, 0, "32768")]
    [InlineData("decimalEnclosedParen", 65535, 0, "65535")]
    [InlineData("decimalEnclosedParen", 999999, 0, "999999")]
    [InlineData("decimalEnclosedParen", 2147483647, 0, "2147483647")]
    [InlineData("direct-decimalEnclosedFullstop", -2147483648, 0, "")]
    [InlineData("direct-decimalEnclosedFullstop", -1, 0, "-1")]
    [InlineData("direct-decimalEnclosedFullstop", 0, 0, "0")]
    [InlineData("direct-decimalEnclosedFullstop", 1, 0, "\u2488")]
    [InlineData("direct-decimalEnclosedFullstop", 9, 0, "\u2490")]
    [InlineData("direct-decimalEnclosedFullstop", 10, 0, "\u2491")]
    [InlineData("direct-decimalEnclosedFullstop", 99, 0, "99")]
    [InlineData("direct-decimalEnclosedFullstop", 100, 0, "100")]
    [InlineData("direct-decimalEnclosedFullstop", 999, 0, "999")]
    [InlineData("direct-decimalEnclosedFullstop", 1000, 0, "1000")]
    [InlineData("direct-decimalEnclosedFullstop", 9999, 0, "9999")]
    [InlineData("direct-decimalEnclosedFullstop", 10000, 0, "10000")]
    [InlineData("direct-decimalEnclosedFullstop", 32767, 0, "32767")]
    [InlineData("direct-decimalEnclosedFullstop", 32768, 0, "32768")]
    [InlineData("direct-decimalEnclosedFullstop", 65535, 0, "65535")]
    [InlineData("direct-decimalEnclosedFullstop", 999999, 0, "999999")]
    [InlineData("direct-decimalEnclosedFullstop", 2147483647, 0, "2147483647")]
    [InlineData("direct-decimalEnclosedCircle", -2147483648, 0, "")]
    [InlineData("direct-decimalEnclosedCircle", -1, 0, "-1")]
    [InlineData("direct-decimalEnclosedCircle", 0, 0, "0")]
    [InlineData("direct-decimalEnclosedCircle", 1, 0, "\u2460")]
    [InlineData("direct-decimalEnclosedCircle", 9, 0, "\u2468")]
    [InlineData("direct-decimalEnclosedCircle", 10, 0, "\u2469")]
    [InlineData("direct-decimalEnclosedCircle", 99, 0, "99")]
    [InlineData("direct-decimalEnclosedCircle", 100, 0, "100")]
    [InlineData("direct-decimalEnclosedCircle", 999, 0, "999")]
    [InlineData("direct-decimalEnclosedCircle", 1000, 0, "1000")]
    [InlineData("direct-decimalEnclosedCircle", 9999, 0, "9999")]
    [InlineData("direct-decimalEnclosedCircle", 10000, 0, "10000")]
    [InlineData("direct-decimalEnclosedCircle", 32767, 0, "32767")]
    [InlineData("direct-decimalEnclosedCircle", 32768, 0, "32768")]
    [InlineData("direct-decimalEnclosedCircle", 65535, 0, "65535")]
    [InlineData("direct-decimalEnclosedCircle", 999999, 0, "999999")]
    [InlineData("direct-decimalEnclosedCircle", 2147483647, 0, "2147483647")]
    [InlineData("direct-decimalEnclosedParen", -2147483648, 0, "")]
    [InlineData("direct-decimalEnclosedParen", -1, 0, "-1")]
    [InlineData("direct-decimalEnclosedParen", 0, 0, "0")]
    [InlineData("direct-decimalEnclosedParen", 1, 0, "\u2474")]
    [InlineData("direct-decimalEnclosedParen", 9, 0, "\u247c")]
    [InlineData("direct-decimalEnclosedParen", 10, 0, "\u247d")]
    [InlineData("direct-decimalEnclosedParen", 99, 0, "99")]
    [InlineData("direct-decimalEnclosedParen", 100, 0, "100")]
    [InlineData("direct-decimalEnclosedParen", 999, 0, "999")]
    [InlineData("direct-decimalEnclosedParen", 1000, 0, "1000")]
    [InlineData("direct-decimalEnclosedParen", 9999, 0, "9999")]
    [InlineData("direct-decimalEnclosedParen", 10000, 0, "10000")]
    [InlineData("direct-decimalEnclosedParen", 32767, 0, "32767")]
    [InlineData("direct-decimalEnclosedParen", 32768, 0, "32768")]
    [InlineData("direct-decimalEnclosedParen", 65535, 0, "65535")]
    [InlineData("direct-decimalEnclosedParen", 999999, 0, "999999")]
    [InlineData("direct-decimalEnclosedParen", 2147483647, 0, "2147483647")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "decimalEnclosedCircle" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u2460, \u2461, \u2462, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "decimalEnclosedFullstop" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u2488, \u2489, \u248a, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "decimalEnclosedParen" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u2474, \u2475, \u2476, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
