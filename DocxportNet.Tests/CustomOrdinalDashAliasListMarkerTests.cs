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

public sealed class CustomOrdinalDashAliasListMarkerTests
{
    // Ordinal and dash aliases and direct formats, checked against Word captures.
    [Theory]
    [InlineData("numberInDash", -2147483648, 0, "")]
    [InlineData("numberInDash", -1, 0, "- -1 -")]
    [InlineData("numberInDash", 0, 0, "- 0 -")]
    [InlineData("numberInDash", 1, 0, "- 1 -")]
    [InlineData("numberInDash", 9, 0, "- 9 -")]
    [InlineData("numberInDash", 10, 0, "- 10 -")]
    [InlineData("numberInDash", 99, 0, "- 99 -")]
    [InlineData("numberInDash", 100, 0, "- 100 -")]
    [InlineData("numberInDash", 999, 0, "- 999 -")]
    [InlineData("numberInDash", 1000, 0, "- 1000 -")]
    [InlineData("numberInDash", 9999, 0, "- 9999 -")]
    [InlineData("numberInDash", 10000, 0, "- 10000 -")]
    [InlineData("numberInDash", 32767, 0, "- 32767 -")]
    [InlineData("numberInDash", 32768, 0, "- 32768 -")]
    [InlineData("numberInDash", 65535, 0, "- 65535 -")]
    [InlineData("numberInDash", 999999, 0, "- 999999 -")]
    [InlineData("numberInDash", 2147483647, 0, "- 2147483647 -")]
    [InlineData("ordinal", -2147483648, 0, "")]
    [InlineData("ordinal", -1, 0, "-1st")]
    [InlineData("ordinal", 0, 0, "0th")]
    [InlineData("ordinal", 1, 0, "1st")]
    [InlineData("ordinal", 9, 0, "9th")]
    [InlineData("ordinal", 10, 0, "10th")]
    [InlineData("ordinal", 99, 0, "99th")]
    [InlineData("ordinal", 100, 0, "100th")]
    [InlineData("ordinal", 999, 0, "999th")]
    [InlineData("ordinal", 1000, 0, "1000th")]
    [InlineData("ordinal", 9999, 0, "9999th")]
    [InlineData("ordinal", 10000, 0, "10000th")]
    [InlineData("ordinal", 32767, 0, "32767th")]
    [InlineData("ordinal", 32768, 0, "32768th")]
    [InlineData("ordinal", 65535, 0, "65535th")]
    [InlineData("ordinal", 999999, 0, "999999th")]
    [InlineData("ordinal", 2147483647, 0, "2147483647th")]
    [InlineData("direct-ordinal", -2147483648, 0, "")]
    [InlineData("direct-ordinal", -1, 0, "-1st")]
    [InlineData("direct-ordinal", 0, 0, "0th")]
    [InlineData("direct-ordinal", 1, 0, "1st")]
    [InlineData("direct-ordinal", 9, 0, "9th")]
    [InlineData("direct-ordinal", 10, 0, "10th")]
    [InlineData("direct-ordinal", 99, 0, "99th")]
    [InlineData("direct-ordinal", 100, 0, "100th")]
    [InlineData("direct-ordinal", 999, 0, "999th")]
    [InlineData("direct-ordinal", 1000, 0, "1000th")]
    [InlineData("direct-ordinal", 9999, 0, "9999th")]
    [InlineData("direct-ordinal", 10000, 0, "10000th")]
    [InlineData("direct-ordinal", 32767, 0, "32767th")]
    [InlineData("direct-ordinal", 32768, 0, "32768th")]
    [InlineData("direct-ordinal", 65535, 0, "65535th")]
    [InlineData("direct-ordinal", 999999, 0, "999999th")]
    [InlineData("direct-ordinal", 2147483647, 0, "2147483647th")]
    [InlineData("direct-numberInDash", -2147483648, 0, "")]
    [InlineData("direct-numberInDash", -1, 0, "- -1 -")]
    [InlineData("direct-numberInDash", 0, 0, "- 0 -")]
    [InlineData("direct-numberInDash", 1, 0, "- 1 -")]
    [InlineData("direct-numberInDash", 9, 0, "- 9 -")]
    [InlineData("direct-numberInDash", 10, 0, "- 10 -")]
    [InlineData("direct-numberInDash", 99, 0, "- 99 -")]
    [InlineData("direct-numberInDash", 100, 0, "- 100 -")]
    [InlineData("direct-numberInDash", 999, 0, "- 999 -")]
    [InlineData("direct-numberInDash", 1000, 0, "- 1000 -")]
    [InlineData("direct-numberInDash", 9999, 0, "- 9999 -")]
    [InlineData("direct-numberInDash", 10000, 0, "- 10000 -")]
    [InlineData("direct-numberInDash", 32767, 0, "- 32767 -")]
    [InlineData("direct-numberInDash", 32768, 0, "- 32768 -")]
    [InlineData("direct-numberInDash", 65535, 0, "- 65535 -")]
    [InlineData("direct-numberInDash", 999999, 0, "- 999999 -")]
    [InlineData("direct-numberInDash", 2147483647, 0, "- 2147483647 -")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "numberInDash" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"- 1 -, - 2 -, - 3 -, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ordinal" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"1st, 2nd, 3rd, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
