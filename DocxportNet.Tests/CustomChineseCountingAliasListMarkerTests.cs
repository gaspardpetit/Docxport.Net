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

public sealed class CustomChineseCountingAliasListMarkerTests
{
    // Chinese and Taiwanese counting aliases and direct formats, checked against Word captures.
    [Theory]
    [InlineData("chineseCounting", -2147483648, 0, "")]
    [InlineData("chineseCounting", -1, 0, "")]
    [InlineData("chineseCounting", 0, 0, "\u25cb")]
    [InlineData("chineseCounting", 1, 0, "\u4e00")]
    [InlineData("chineseCounting", 9, 0, "\u4e5d")]
    [InlineData("chineseCounting", 10, 0, "\u5341")]
    [InlineData("chineseCounting", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("chineseCounting", 100, 0, "\u4e00\u25cb\u25cb")]
    [InlineData("chineseCounting", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("chineseCounting", 1000, 0, "\u4e00\u25cb\u25cb\u25cb")]
    [InlineData("chineseCounting", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("chineseCounting", 10000, 0, "\u4e00\u25cb\u25cb\u25cb\u25cb")]
    [InlineData("chineseCounting", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("chineseCounting", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("chineseCounting", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("chineseCounting", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("chineseCounting", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    [InlineData("chineseCountingThousand", -2147483648, 0, "")]
    [InlineData("chineseCountingThousand", -1, 0, "")]
    [InlineData("chineseCountingThousand", 0, 0, "\u3007")]
    [InlineData("chineseCountingThousand", 1, 0, "\u4e00")]
    [InlineData("chineseCountingThousand", 9, 0, "\u4e5d")]
    [InlineData("chineseCountingThousand", 10, 0, "\u5341")]
    [InlineData("chineseCountingThousand", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("chineseCountingThousand", 100, 0, "\u4e00\u767e")]
    [InlineData("chineseCountingThousand", 999, 0, "\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("chineseCountingThousand", 1000, 0, "\u4e00\u5343")]
    [InlineData("chineseCountingThousand", 9999, 0, "\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("chineseCountingThousand", 10000, 0, "\u4e00\u4e07")]
    [InlineData("chineseCountingThousand", 32767, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u4e03")]
    [InlineData("chineseCountingThousand", 32768, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u516b")]
    [InlineData("chineseCountingThousand", 65535, 0, "\u516d\u4e07\u4e94\u5343\u4e94\u767e\u4e09\u5341\u4e94")]
    [InlineData("chineseCountingThousand", 999999, 0, "\u4e5d\u5341\u4e5d\u4e07\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("chineseCountingThousand", 2147483647, 0, "")]
    [InlineData("taiwaneseCounting", -2147483648, 0, "")]
    [InlineData("taiwaneseCounting", -1, 0, "")]
    [InlineData("taiwaneseCounting", 0, 0, "\u25cb")]
    [InlineData("taiwaneseCounting", 1, 0, "\u4e00")]
    [InlineData("taiwaneseCounting", 9, 0, "\u4e5d")]
    [InlineData("taiwaneseCounting", 10, 0, "\u5341")]
    [InlineData("taiwaneseCounting", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("taiwaneseCounting", 100, 0, "\u4e00\u25cb\u25cb")]
    [InlineData("taiwaneseCounting", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseCounting", 1000, 0, "\u4e00\u25cb\u25cb\u25cb")]
    [InlineData("taiwaneseCounting", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseCounting", 10000, 0, "\u4e00\u25cb\u25cb\u25cb\u25cb")]
    [InlineData("taiwaneseCounting", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("taiwaneseCounting", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("taiwaneseCounting", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("taiwaneseCounting", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseCounting", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    [InlineData("taiwaneseCountingThousand", -2147483648, 0, "")]
    [InlineData("taiwaneseCountingThousand", -1, 0, "")]
    [InlineData("taiwaneseCountingThousand", 0, 0, "\u96f6")]
    [InlineData("taiwaneseCountingThousand", 1, 0, "\u4e00")]
    [InlineData("taiwaneseCountingThousand", 9, 0, "\u4e5d")]
    [InlineData("taiwaneseCountingThousand", 10, 0, "\u5341")]
    [InlineData("taiwaneseCountingThousand", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("taiwaneseCountingThousand", 100, 0, "\u4e00\u767e")]
    [InlineData("taiwaneseCountingThousand", 999, 0, "\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("taiwaneseCountingThousand", 1000, 0, "\u4e00\u5343")]
    [InlineData("taiwaneseCountingThousand", 9999, 0, "\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("taiwaneseCountingThousand", 10000, 0, "\u4e00\u842c")]
    [InlineData("taiwaneseCountingThousand", 32767, 0, "\u4e09\u842c\u4e8c\u5343\u4e03\u767e\u516d\u5341\u4e03")]
    [InlineData("taiwaneseCountingThousand", 32768, 0, "\u4e09\u842c\u4e8c\u5343\u4e03\u767e\u516d\u5341\u516b")]
    [InlineData("taiwaneseCountingThousand", 65535, 0, "\u516d\u842c\u4e94\u5343\u4e94\u767e\u4e09\u5341\u4e94")]
    [InlineData("taiwaneseCountingThousand", 999999, 0, "\u4e5d\u5341\u4e5d\u842c\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("taiwaneseCountingThousand", 2147483647, 0, "")]
    [InlineData("direct-chineseCounting", -2147483648, 0, "")]
    [InlineData("direct-chineseCounting", -1, 0, "")]
    [InlineData("direct-chineseCounting", 0, 0, "\u25cb")]
    [InlineData("direct-chineseCounting", 1, 0, "\u4e00")]
    [InlineData("direct-chineseCounting", 9, 0, "\u4e5d")]
    [InlineData("direct-chineseCounting", 10, 0, "\u5341")]
    [InlineData("direct-chineseCounting", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("direct-chineseCounting", 100, 0, "\u4e00\u25cb\u25cb")]
    [InlineData("direct-chineseCounting", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-chineseCounting", 1000, 0, "\u4e00\u25cb\u25cb\u25cb")]
    [InlineData("direct-chineseCounting", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-chineseCounting", 10000, 0, "\u4e00\u25cb\u25cb\u25cb\u25cb")]
    [InlineData("direct-chineseCounting", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("direct-chineseCounting", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("direct-chineseCounting", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("direct-chineseCounting", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-chineseCounting", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    [InlineData("direct-chineseCountingThousand", -2147483648, 0, "")]
    [InlineData("direct-chineseCountingThousand", -1, 0, "")]
    [InlineData("direct-chineseCountingThousand", 0, 0, "\u3007")]
    [InlineData("direct-chineseCountingThousand", 1, 0, "\u4e00")]
    [InlineData("direct-chineseCountingThousand", 9, 0, "\u4e5d")]
    [InlineData("direct-chineseCountingThousand", 10, 0, "\u5341")]
    [InlineData("direct-chineseCountingThousand", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("direct-chineseCountingThousand", 100, 0, "\u4e00\u767e")]
    [InlineData("direct-chineseCountingThousand", 999, 0, "\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-chineseCountingThousand", 1000, 0, "\u4e00\u5343")]
    [InlineData("direct-chineseCountingThousand", 9999, 0, "\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-chineseCountingThousand", 10000, 0, "\u4e00\u4e07")]
    [InlineData("direct-chineseCountingThousand", 32767, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u4e03")]
    [InlineData("direct-chineseCountingThousand", 32768, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u516b")]
    [InlineData("direct-chineseCountingThousand", 65535, 0, "\u516d\u4e07\u4e94\u5343\u4e94\u767e\u4e09\u5341\u4e94")]
    [InlineData("direct-chineseCountingThousand", 999999, 0, "\u4e5d\u5341\u4e5d\u4e07\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-chineseCountingThousand", 2147483647, 0, "")]
    [InlineData("direct-taiwaneseCounting", -2147483648, 0, "")]
    [InlineData("direct-taiwaneseCounting", -1, 0, "")]
    [InlineData("direct-taiwaneseCounting", 0, 0, "\u25cb")]
    [InlineData("direct-taiwaneseCounting", 1, 0, "\u4e00")]
    [InlineData("direct-taiwaneseCounting", 9, 0, "\u4e5d")]
    [InlineData("direct-taiwaneseCounting", 10, 0, "\u5341")]
    [InlineData("direct-taiwaneseCounting", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("direct-taiwaneseCounting", 100, 0, "\u4e00\u25cb\u25cb")]
    [InlineData("direct-taiwaneseCounting", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-taiwaneseCounting", 1000, 0, "\u4e00\u25cb\u25cb\u25cb")]
    [InlineData("direct-taiwaneseCounting", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-taiwaneseCounting", 10000, 0, "\u4e00\u25cb\u25cb\u25cb\u25cb")]
    [InlineData("direct-taiwaneseCounting", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("direct-taiwaneseCounting", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("direct-taiwaneseCounting", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("direct-taiwaneseCounting", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-taiwaneseCounting", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    [InlineData("direct-taiwaneseCountingThousand", -2147483648, 0, "")]
    [InlineData("direct-taiwaneseCountingThousand", -1, 0, "")]
    [InlineData("direct-taiwaneseCountingThousand", 0, 0, "\u96f6")]
    [InlineData("direct-taiwaneseCountingThousand", 1, 0, "\u4e00")]
    [InlineData("direct-taiwaneseCountingThousand", 9, 0, "\u4e5d")]
    [InlineData("direct-taiwaneseCountingThousand", 10, 0, "\u5341")]
    [InlineData("direct-taiwaneseCountingThousand", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("direct-taiwaneseCountingThousand", 100, 0, "\u4e00\u767e")]
    [InlineData("direct-taiwaneseCountingThousand", 999, 0, "\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-taiwaneseCountingThousand", 1000, 0, "\u4e00\u5343")]
    [InlineData("direct-taiwaneseCountingThousand", 9999, 0, "\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-taiwaneseCountingThousand", 10000, 0, "\u4e00\u842c")]
    [InlineData("direct-taiwaneseCountingThousand", 32767, 0, "\u4e09\u842c\u4e8c\u5343\u4e03\u767e\u516d\u5341\u4e03")]
    [InlineData("direct-taiwaneseCountingThousand", 32768, 0, "\u4e09\u842c\u4e8c\u5343\u4e03\u767e\u516d\u5341\u516b")]
    [InlineData("direct-taiwaneseCountingThousand", 65535, 0, "\u516d\u842c\u4e94\u5343\u4e94\u767e\u4e09\u5341\u4e94")]
    [InlineData("direct-taiwaneseCountingThousand", 999999, 0, "\u4e5d\u5341\u4e5d\u842c\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-taiwaneseCountingThousand", 2147483647, 0, "")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "chineseCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u5341, \u4e00\u25cb\u25cb(\u7b80), ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "chineseCountingThousand" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u5341, \u4e00\u767e(\u7b80), ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "taiwaneseCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u5341, \u4e00\u25cb\u25cb(\u7e41), ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "taiwaneseCountingThousand" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u5341, \u4e00\u767e(\u7e41), ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
