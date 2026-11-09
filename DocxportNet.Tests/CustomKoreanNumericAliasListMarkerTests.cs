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

public sealed class CustomKoreanNumericAliasListMarkerTests
{
    // Word captures for Korean counting and digit custom aliases.
    [Theory]
    [InlineData("koreanCounting", -2147483648, 0, "")]
    [InlineData("koreanCounting", -1, 0, "")]
    [InlineData("koreanCounting", 0, 0, "\uc601")]
    [InlineData("koreanCounting", 1, 0, "\uc77c")]
    [InlineData("koreanCounting", 9, 0, "\uad6c")]
    [InlineData("koreanCounting", 10, 0, "\uc2ed")]
    [InlineData("koreanCounting", 99, 0, "\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 100, 0, "\ubc31")]
    [InlineData("koreanCounting", 999, 0, "\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 1000, 0, "\ucc9c")]
    [InlineData("koreanCounting", 9999, 0, "\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 10000, 0, "\ub9cc")]
    [InlineData("koreanCounting", 32767, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\uce60")]
    [InlineData("koreanCounting", 32768, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\ud314")]
    [InlineData("koreanCounting", 65535, 0, "\uc721\ub9cc\uc624\ucc9c\uc624\ubc31\uc0bc\uc2ed\uc624")]
    [InlineData("koreanCounting", 999999, 0, "\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanCounting", 2147483647, 0, "")]
    [InlineData("koreanDigital", -2147483648, 0, "")]
    [InlineData("koreanDigital", -1, 0, "")]
    [InlineData("koreanDigital", 0, 0, "\uc601")]
    [InlineData("koreanDigital", 1, 0, "\uc77c")]
    [InlineData("koreanDigital", 9, 0, "\uad6c")]
    [InlineData("koreanDigital", 10, 0, "\uc77c\uc601")]
    [InlineData("koreanDigital", 99, 0, "\uad6c\uad6c")]
    [InlineData("koreanDigital", 100, 0, "\uc77c\uc601\uc601")]
    [InlineData("koreanDigital", 999, 0, "\uad6c\uad6c\uad6c")]
    [InlineData("koreanDigital", 1000, 0, "\uc77c\uc601\uc601\uc601")]
    [InlineData("koreanDigital", 9999, 0, "\uad6c\uad6c\uad6c\uad6c")]
    [InlineData("koreanDigital", 10000, 0, "\uc77c\uc601\uc601\uc601\uc601")]
    [InlineData("koreanDigital", 32767, 0, "\uc0bc\uc774\uce60\uc721\uce60")]
    [InlineData("koreanDigital", 32768, 0, "\uc0bc\uc774\uce60\uc721\ud314")]
    [InlineData("koreanDigital", 65535, 0, "\uc721\uc624\uc624\uc0bc\uc624")]
    [InlineData("koreanDigital", 999999, 0, "\uad6c\uad6c\uad6c\uad6c\uad6c\uad6c")]
    [InlineData("koreanDigital", 2147483647, 0, "\uc774\uc77c\uc0ac\uce60\uc0ac\ud314\uc0bc\uc721\uc0ac\uce60")]
    [InlineData("koreanDigital2", -2147483648, 0, "")]
    [InlineData("koreanDigital2", -1, 0, "")]
    [InlineData("koreanDigital2", 0, 0, "\uf9b2")]
    [InlineData("koreanDigital2", 1, 0, "\u4e00")]
    [InlineData("koreanDigital2", 9, 0, "\u4e5d")]
    [InlineData("koreanDigital2", 10, 0, "\u4e00\u96f6")]
    [InlineData("koreanDigital2", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("koreanDigital2", 100, 0, "\u4e00\u96f6\u96f6")]
    [InlineData("koreanDigital2", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("koreanDigital2", 1000, 0, "\u4e00\u96f6\u96f6\u96f6")]
    [InlineData("koreanDigital2", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("koreanDigital2", 10000, 0, "\u4e00\u96f6\u96f6\u96f6\u96f6")]
    [InlineData("koreanDigital2", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("koreanDigital2", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("koreanDigital2", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("koreanDigital2", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("koreanDigital2", 2147483647, 0, "")]
    [InlineData("koreanLegal", -2147483648, 0, "")]
    [InlineData("koreanLegal", -1, 0, "-1")]
    [InlineData("koreanLegal", 0, 0, "0")]
    [InlineData("koreanLegal", 1, 0, "\ud558\ub098")]
    [InlineData("koreanLegal", 9, 0, "\uc544\ud649")]
    [InlineData("koreanLegal", 10, 0, "\uc5f4")]
    [InlineData("koreanLegal", 99, 0, "\uc544\ud754\uc544\ud649")]
    [InlineData("koreanLegal", 100, 0, "\ubc31")]
    [InlineData("koreanLegal", 999, 0, "\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanLegal", 1000, 0, "\ucc9c")]
    [InlineData("koreanLegal", 9999, 0, "\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanLegal", 10000, 0, "\ub9cc")]
    [InlineData("koreanLegal", 32767, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\uce60")]
    [InlineData("koreanLegal", 32768, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\ud314")]
    [InlineData("koreanLegal", 65535, 0, "\uc721\ub9cc\uc624\ucc9c\uc624\ubc31\uc0bc\uc2ed\uc624")]
    [InlineData("koreanLegal", 999999, 0, "\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("koreanLegal", 2147483647, 0, "")]
    [InlineData("direct-koreanCounting", -2147483648, 0, "")]
    [InlineData("direct-koreanCounting", -1, 0, "")]
    [InlineData("direct-koreanCounting", 0, 0, "\uc601")]
    [InlineData("direct-koreanCounting", 1, 0, "\uc77c")]
    [InlineData("direct-koreanCounting", 9, 0, "\uad6c")]
    [InlineData("direct-koreanCounting", 10, 0, "\uc2ed")]
    [InlineData("direct-koreanCounting", 99, 0, "\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanCounting", 100, 0, "\ubc31")]
    [InlineData("direct-koreanCounting", 999, 0, "\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanCounting", 1000, 0, "\ucc9c")]
    [InlineData("direct-koreanCounting", 9999, 0, "\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanCounting", 10000, 0, "\ub9cc")]
    [InlineData("direct-koreanCounting", 32767, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\uce60")]
    [InlineData("direct-koreanCounting", 32768, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\ud314")]
    [InlineData("direct-koreanCounting", 65535, 0, "\uc721\ub9cc\uc624\ucc9c\uc624\ubc31\uc0bc\uc2ed\uc624")]
    [InlineData("direct-koreanCounting", 999999, 0, "\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanCounting", 2147483647, 0, "")]
    [InlineData("direct-koreanDigital2", -2147483648, 0, "")]
    [InlineData("direct-koreanDigital2", -1, 0, "")]
    [InlineData("direct-koreanDigital2", 0, 0, "\uf9b2")]
    [InlineData("direct-koreanDigital2", 1, 0, "\u4e00")]
    [InlineData("direct-koreanDigital2", 9, 0, "\u4e5d")]
    [InlineData("direct-koreanDigital2", 10, 0, "\u4e00\u96f6")]
    [InlineData("direct-koreanDigital2", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("direct-koreanDigital2", 100, 0, "\u4e00\u96f6\u96f6")]
    [InlineData("direct-koreanDigital2", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-koreanDigital2", 1000, 0, "\u4e00\u96f6\u96f6\u96f6")]
    [InlineData("direct-koreanDigital2", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-koreanDigital2", 10000, 0, "\u4e00\u96f6\u96f6\u96f6\u96f6")]
    [InlineData("direct-koreanDigital2", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("direct-koreanDigital2", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("direct-koreanDigital2", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("direct-koreanDigital2", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("direct-koreanDigital2", 2147483647, 0, "")]
    [InlineData("direct-koreanLegal", -2147483648, 0, "")]
    [InlineData("direct-koreanLegal", -1, 0, "-1")]
    [InlineData("direct-koreanLegal", 0, 0, "0")]
    [InlineData("direct-koreanLegal", 1, 0, "\ud558\ub098")]
    [InlineData("direct-koreanLegal", 9, 0, "\uc544\ud649")]
    [InlineData("direct-koreanLegal", 10, 0, "\uc5f4")]
    [InlineData("direct-koreanLegal", 99, 0, "\uc544\ud754\uc544\ud649")]
    [InlineData("direct-koreanLegal", 100, 0, "\ubc31")]
    [InlineData("direct-koreanLegal", 999, 0, "\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanLegal", 1000, 0, "\ucc9c")]
    [InlineData("direct-koreanLegal", 9999, 0, "\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanLegal", 10000, 0, "\ub9cc")]
    [InlineData("direct-koreanLegal", 32767, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\uce60")]
    [InlineData("direct-koreanLegal", 32768, 0, "\uc0bc\ub9cc\uc774\ucc9c\uce60\ubc31\uc721\uc2ed\ud314")]
    [InlineData("direct-koreanLegal", 65535, 0, "\uc721\ub9cc\uc624\ucc9c\uc624\ubc31\uc0bc\uc2ed\uc624")]
    [InlineData("direct-koreanLegal", 999999, 0, "\uad6c\uc2ed\uad6c\ub9cc\uad6c\ucc9c\uad6c\ubc31\uad6c\uc2ed\uad6c")]
    [InlineData("direct-koreanLegal", 2147483647, 0, "")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "koreanCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\uc77c, \uc774, \uc0bc, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "koreanDigital" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\uc77c, \uc77c\uc601, \uc77c\uc601\uc601, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "koreanDigital2" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u4e00\u96f6, \u4e00\u96f6\u96f6, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "koreanLegal" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\ud558\ub098, \ub458, \uc14b, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abstractNumbering = new AbstractNum(abstractXml);
        if (directFormat != null)
            abstractNumbering.GetFirstChild<Level>()!.NumberingFormat = new NumberingFormat { Val = new NumberFormatValues(directFormat) };
        var numbering = new Numbering(abstractNumbering,
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
