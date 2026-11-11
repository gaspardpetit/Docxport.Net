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

public sealed class CustomJapaneseAliasListMarkerTests
{
    // Japanese numeric aliases and direct formats, checked against Word captures.
    [Theory]
    [InlineData("japaneseCounting", -2147483648, 0, "")]
    [InlineData("japaneseCounting", -1, 0, "")]
    [InlineData("japaneseCounting", 0, 0, "\u3007")]
    [InlineData("japaneseCounting", 1, 0, "\u4e00")]
    [InlineData("japaneseCounting", 9, 0, "\u4e5d")]
    [InlineData("japaneseCounting", 10, 0, "\u5341")]
    [InlineData("japaneseCounting", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("japaneseCounting", 100, 0, "\u767e")]
    [InlineData("japaneseCounting", 999, 0, "\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("japaneseCounting", 1000, 0, "\u5343")]
    [InlineData("japaneseCounting", 9999, 0, "\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("japaneseCounting", 10000, 0, "\u4e00\u4e07")]
    [InlineData("japaneseCounting", 32767, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u4e03")]
    [InlineData("japaneseCounting", 32768, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u516b")]
    [InlineData("japaneseCounting", 65535, 0, "\u516d\u4e07\u4e94\u5343\u4e94\u767e\u4e09\u5341\u4e94")]
    [InlineData("japaneseCounting", 999999, 0, "\u4e5d\u5341\u4e5d\u4e07\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("japaneseCounting", 2147483647, 0, "")]
    [InlineData("japaneseDigitalTenThousand", -2147483648, 0, "")]
    [InlineData("japaneseDigitalTenThousand", -1, 0, "")]
    [InlineData("japaneseDigitalTenThousand", 0, 0, "\u3007")]
    [InlineData("japaneseDigitalTenThousand", 1, 0, "\u4e00")]
    [InlineData("japaneseDigitalTenThousand", 9, 0, "\u4e5d")]
    [InlineData("japaneseDigitalTenThousand", 10, 0, "\u4e00\u3007")]
    [InlineData("japaneseDigitalTenThousand", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("japaneseDigitalTenThousand", 100, 0, "\u4e00\u3007\u3007")]
    [InlineData("japaneseDigitalTenThousand", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("japaneseDigitalTenThousand", 1000, 0, "\u4e00\u3007\u3007\u3007")]
    [InlineData("japaneseDigitalTenThousand", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("japaneseDigitalTenThousand", 10000, 0, "")]
    [InlineData("japaneseDigitalTenThousand", 32767, 0, "")]
    [InlineData("japaneseDigitalTenThousand", 32768, 0, "")]
    [InlineData("japaneseDigitalTenThousand", 65535, 0, "")]
    [InlineData("japaneseDigitalTenThousand", 999999, 0, "")]
    [InlineData("japaneseDigitalTenThousand", 2147483647, 0, "")]
    [InlineData("japaneseLegal", -2147483648, 0, "")]
    [InlineData("japaneseLegal", -1, 0, "")]
    [InlineData("japaneseLegal", 0, 0, "\u3007")]
    [InlineData("japaneseLegal", 1, 0, "\u58f1")]
    [InlineData("japaneseLegal", 9, 0, "\u4e5d")]
    [InlineData("japaneseLegal", 10, 0, "\u58f1\u62fe")]
    [InlineData("japaneseLegal", 99, 0, "\u4e5d\u62fe\u4e5d")]
    [InlineData("japaneseLegal", 100, 0, "\u58f1\u767e")]
    [InlineData("japaneseLegal", 999, 0, "\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("japaneseLegal", 1000, 0, "\u58f1\u9621")]
    [InlineData("japaneseLegal", 9999, 0, "\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("japaneseLegal", 10000, 0, "\u58f1\u842c")]
    [InlineData("japaneseLegal", 32767, 0, "\u53c2\u842c\u5f10\u9621\u4e03\u767e\u516d\u62fe\u4e03")]
    [InlineData("japaneseLegal", 32768, 0, "\u53c2\u842c\u5f10\u9621\u4e03\u767e\u516d\u62fe\u516b")]
    [InlineData("japaneseLegal", 65535, 0, "\u516d\u842c\u4f0d\u9621\u4f0d\u767e\u53c2\u62fe\u4f0d")]
    [InlineData("japaneseLegal", 999999, 0, "\u4e5d\u62fe\u4e5d\u842c\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("japaneseLegal", 2147483647, 0, "\u5f10\u62fe\u58f1\u5104\u56db\u9621\u4e03\u767e\u56db\u62fe\u516b\u842c\u53c2\u9621\u516d\u767e\u56db\u62fe\u4e03")]
    [InlineData("direct-japaneseCounting", -2147483648, 0, "")]
    [InlineData("direct-japaneseCounting", -1, 0, "")]
    [InlineData("direct-japaneseCounting", 0, 0, "\u3007")]
    [InlineData("direct-japaneseCounting", 1, 0, "\u4e00")]
    [InlineData("direct-japaneseCounting", 9, 0, "\u4e5d")]
    [InlineData("direct-japaneseCounting", 10, 0, "\u5341")]
    [InlineData("direct-japaneseCounting", 99, 0, "\u4e5d\u5341\u4e5d")]
    [InlineData("direct-japaneseCounting", 100, 0, "\u767e")]
    [InlineData("direct-japaneseCounting", 999, 0, "\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-japaneseCounting", 1000, 0, "\u5343")]
    [InlineData("direct-japaneseCounting", 9999, 0, "\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-japaneseCounting", 10000, 0, "\u4e00\u4e07")]
    [InlineData("direct-japaneseCounting", 32767, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u4e03")]
    [InlineData("direct-japaneseCounting", 32768, 0, "\u4e09\u4e07\u4e8c\u5343\u4e03\u767e\u516d\u5341\u516b")]
    [InlineData("direct-japaneseCounting", 65535, 0, "\u516d\u4e07\u4e94\u5343\u4e94\u767e\u4e09\u5341\u4e94")]
    [InlineData("direct-japaneseCounting", 999999, 0, "\u4e5d\u5341\u4e5d\u4e07\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5d")]
    [InlineData("direct-japaneseCounting", 2147483647, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", -2147483648, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", -1, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", 0, 0, "\u3007")]
    [InlineData("direct-japaneseDigitalTenThousand", 1, 0, "\u4e00")]
    [InlineData("direct-japaneseDigitalTenThousand", 9, 0, "\u4e5d")]
    [InlineData("direct-japaneseDigitalTenThousand", 10, 0, "\u4e00\u3007")]
    [InlineData("direct-japaneseDigitalTenThousand", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("direct-japaneseDigitalTenThousand", 100, 0, "\u4e00\u3007\u3007")]
    [InlineData("direct-japaneseDigitalTenThousand", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-japaneseDigitalTenThousand", 1000, 0, "\u4e00\u3007\u3007\u3007")]
    [InlineData("direct-japaneseDigitalTenThousand", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-japaneseDigitalTenThousand", 10000, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", 32767, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", 32768, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", 65535, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", 999999, 0, "")]
    [InlineData("direct-japaneseDigitalTenThousand", 2147483647, 0, "")]
    [InlineData("direct-japaneseLegal", -2147483648, 0, "")]
    [InlineData("direct-japaneseLegal", -1, 0, "")]
    [InlineData("direct-japaneseLegal", 0, 0, "\u3007")]
    [InlineData("direct-japaneseLegal", 1, 0, "\u58f1")]
    [InlineData("direct-japaneseLegal", 9, 0, "\u4e5d")]
    [InlineData("direct-japaneseLegal", 10, 0, "\u58f1\u62fe")]
    [InlineData("direct-japaneseLegal", 99, 0, "\u4e5d\u62fe\u4e5d")]
    [InlineData("direct-japaneseLegal", 100, 0, "\u58f1\u767e")]
    [InlineData("direct-japaneseLegal", 999, 0, "\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("direct-japaneseLegal", 1000, 0, "\u58f1\u9621")]
    [InlineData("direct-japaneseLegal", 9999, 0, "\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("direct-japaneseLegal", 10000, 0, "\u58f1\u842c")]
    [InlineData("direct-japaneseLegal", 32767, 0, "\u53c2\u842c\u5f10\u9621\u4e03\u767e\u516d\u62fe\u4e03")]
    [InlineData("direct-japaneseLegal", 32768, 0, "\u53c2\u842c\u5f10\u9621\u4e03\u767e\u516d\u62fe\u516b")]
    [InlineData("direct-japaneseLegal", 65535, 0, "\u516d\u842c\u4f0d\u9621\u4f0d\u767e\u53c2\u62fe\u4f0d")]
    [InlineData("direct-japaneseLegal", 999999, 0, "\u4e5d\u62fe\u4e5d\u842c\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u4e5d")]
    [InlineData("direct-japaneseLegal", 2147483647, 0, "\u5f10\u62fe\u58f1\u5104\u56db\u9621\u4e03\u767e\u56db\u62fe\u516b\u842c\u53c2\u9621\u516d\u767e\u56db\u62fe\u4e03")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "japaneseCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u4e8c, \u4e09, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "japaneseDigitalTenThousand" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u4e8c, \u4e09, \u4e07, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "japaneseLegal" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u58f1, \u5f10, \u53c2, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
