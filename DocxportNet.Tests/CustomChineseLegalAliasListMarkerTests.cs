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

public sealed class CustomChineseLegalAliasListMarkerTests
{
    // Chinese legal numeral aliases and direct formats, checked against Word captures.
    [Theory]
    [InlineData("chineseLegalSimplified", -2147483648, 0, "")]
    [InlineData("chineseLegalSimplified", -1, 0, "")]
    [InlineData("chineseLegalSimplified", 0, 0, "\u96f6")]
    [InlineData("chineseLegalSimplified", 1, 0, "\u58f9")]
    [InlineData("chineseLegalSimplified", 9, 0, "\u7396")]
    [InlineData("chineseLegalSimplified", 10, 0, "\u58f9\u62fe")]
    [InlineData("chineseLegalSimplified", 99, 0, "\u7396\u62fe\u7396")]
    [InlineData("chineseLegalSimplified", 100, 0, "\u58f9\u4f70")]
    [InlineData("chineseLegalSimplified", 999, 0, "\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("chineseLegalSimplified", 1000, 0, "\u58f9\u4edf")]
    [InlineData("chineseLegalSimplified", 9999, 0, "\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("chineseLegalSimplified", 10000, 0, "\u58f9\u842c")]
    [InlineData("chineseLegalSimplified", 32767, 0, "\u53c1\u842c\u8d30\u4edf\u67d2\u4f70\u9646\u62fe\u67d2")]
    [InlineData("chineseLegalSimplified", 32768, 0, "\u53c1\u842c\u8d30\u4edf\u67d2\u4f70\u9646\u62fe\u634c")]
    [InlineData("chineseLegalSimplified", 65535, 0, "\u9646\u842c\u4f0d\u4edf\u4f0d\u4f70\u53c1\u62fe\u4f0d")]
    [InlineData("chineseLegalSimplified", 999999, 0, "\u7396\u62fe\u7396\u842c\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("chineseLegalSimplified", 2147483647, 0, "")]
    [InlineData("ideographLegalTraditional", -2147483648, 0, "")]
    [InlineData("ideographLegalTraditional", -1, 0, "")]
    [InlineData("ideographLegalTraditional", 0, 0, "\u96f6")]
    [InlineData("ideographLegalTraditional", 1, 0, "\u58f9")]
    [InlineData("ideographLegalTraditional", 9, 0, "\u7396")]
    [InlineData("ideographLegalTraditional", 10, 0, "\u58f9\u62fe")]
    [InlineData("ideographLegalTraditional", 99, 0, "\u7396\u62fe\u7396")]
    [InlineData("ideographLegalTraditional", 100, 0, "\u58f9\u4f70")]
    [InlineData("ideographLegalTraditional", 999, 0, "\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("ideographLegalTraditional", 1000, 0, "\u58f9\u4edf")]
    [InlineData("ideographLegalTraditional", 9999, 0, "\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("ideographLegalTraditional", 10000, 0, "\u58f9\u842c")]
    [InlineData("ideographLegalTraditional", 32767, 0, "\u53c3\u842c\u8cb3\u4edf\u67d2\u4f70\u9678\u62fe\u67d2")]
    [InlineData("ideographLegalTraditional", 32768, 0, "\u53c3\u842c\u8cb3\u4edf\u67d2\u4f70\u9678\u62fe\u634c")]
    [InlineData("ideographLegalTraditional", 65535, 0, "\u9678\u842c\u4f0d\u4edf\u4f0d\u4f70\u53c3\u62fe\u4f0d")]
    [InlineData("ideographLegalTraditional", 999999, 0, "\u7396\u62fe\u7396\u842c\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("ideographLegalTraditional", 2147483647, 0, "")]
    [InlineData("direct-chineseLegalSimplified", -2147483648, 0, "")]
    [InlineData("direct-chineseLegalSimplified", -1, 0, "")]
    [InlineData("direct-chineseLegalSimplified", 0, 0, "\u96f6")]
    [InlineData("direct-chineseLegalSimplified", 1, 0, "\u58f9")]
    [InlineData("direct-chineseLegalSimplified", 9, 0, "\u7396")]
    [InlineData("direct-chineseLegalSimplified", 10, 0, "\u58f9\u62fe")]
    [InlineData("direct-chineseLegalSimplified", 99, 0, "\u7396\u62fe\u7396")]
    [InlineData("direct-chineseLegalSimplified", 100, 0, "\u58f9\u4f70")]
    [InlineData("direct-chineseLegalSimplified", 999, 0, "\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("direct-chineseLegalSimplified", 1000, 0, "\u58f9\u4edf")]
    [InlineData("direct-chineseLegalSimplified", 9999, 0, "\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("direct-chineseLegalSimplified", 10000, 0, "\u58f9\u842c")]
    [InlineData("direct-chineseLegalSimplified", 32767, 0, "\u53c1\u842c\u8d30\u4edf\u67d2\u4f70\u9646\u62fe\u67d2")]
    [InlineData("direct-chineseLegalSimplified", 32768, 0, "\u53c1\u842c\u8d30\u4edf\u67d2\u4f70\u9646\u62fe\u634c")]
    [InlineData("direct-chineseLegalSimplified", 65535, 0, "\u9646\u842c\u4f0d\u4edf\u4f0d\u4f70\u53c1\u62fe\u4f0d")]
    [InlineData("direct-chineseLegalSimplified", 999999, 0, "\u7396\u62fe\u7396\u842c\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("direct-chineseLegalSimplified", 2147483647, 0, "")]
    [InlineData("direct-ideographLegalTraditional", -2147483648, 0, "")]
    [InlineData("direct-ideographLegalTraditional", -1, 0, "")]
    [InlineData("direct-ideographLegalTraditional", 0, 0, "\u96f6")]
    [InlineData("direct-ideographLegalTraditional", 1, 0, "\u58f9")]
    [InlineData("direct-ideographLegalTraditional", 9, 0, "\u7396")]
    [InlineData("direct-ideographLegalTraditional", 10, 0, "\u58f9\u62fe")]
    [InlineData("direct-ideographLegalTraditional", 99, 0, "\u7396\u62fe\u7396")]
    [InlineData("direct-ideographLegalTraditional", 100, 0, "\u58f9\u4f70")]
    [InlineData("direct-ideographLegalTraditional", 999, 0, "\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("direct-ideographLegalTraditional", 1000, 0, "\u58f9\u4edf")]
    [InlineData("direct-ideographLegalTraditional", 9999, 0, "\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("direct-ideographLegalTraditional", 10000, 0, "\u58f9\u842c")]
    [InlineData("direct-ideographLegalTraditional", 32767, 0, "\u53c3\u842c\u8cb3\u4edf\u67d2\u4f70\u9678\u62fe\u67d2")]
    [InlineData("direct-ideographLegalTraditional", 32768, 0, "\u53c3\u842c\u8cb3\u4edf\u67d2\u4f70\u9678\u62fe\u634c")]
    [InlineData("direct-ideographLegalTraditional", 65535, 0, "\u9678\u842c\u4f0d\u4edf\u4f0d\u4f70\u53c3\u62fe\u4f0d")]
    [InlineData("direct-ideographLegalTraditional", 999999, 0, "\u7396\u62fe\u7396\u842c\u7396\u4edf\u7396\u4f70\u7396\u62fe\u7396")]
    [InlineData("direct-ideographLegalTraditional", 2147483647, 0, "")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "chineseLegalSimplified" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u58f9, \u8d30, \u53c1, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ideographLegalTraditional" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u58f9, \u8cb3, \u53c3, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
