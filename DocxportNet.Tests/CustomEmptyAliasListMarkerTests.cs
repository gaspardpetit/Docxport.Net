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

public sealed class CustomEmptyAliasListMarkerTests
{
    // Empty custom alias and direct none controls, checked against fresh Word captures.
    [Theory]
    [InlineData("control-none", -2147483648, 0, "")]
    [InlineData("control-none", -1, 0, "")]
    [InlineData("control-none", 0, 0, "")]
    [InlineData("control-none", 1, 0, "")]
    [InlineData("control-none", 9, 0, "")]
    [InlineData("control-none", 10, 0, "")]
    [InlineData("control-none", 99, 0, "")]
    [InlineData("control-none", 100, 0, "")]
    [InlineData("control-none", 999, 0, "")]
    [InlineData("control-none", 1000, 0, "")]
    [InlineData("control-none", 9999, 0, "")]
    [InlineData("control-none", 10000, 0, "")]
    [InlineData("control-none", 32767, 0, "")]
    [InlineData("control-none", 32768, 0, "")]
    [InlineData("control-none", 65535, 0, "")]
    [InlineData("control-none", 999999, 0, "")]
    [InlineData("control-none", 2147483647, 0, "")]
    [InlineData("none", -2147483648, 0, "")]
    [InlineData("none", -1, 0, "")]
    [InlineData("none", 0, 0, "")]
    [InlineData("none", 1, 0, "")]
    [InlineData("none", 9, 0, "")]
    [InlineData("none", 10, 0, "")]
    [InlineData("none", 99, 0, "")]
    [InlineData("none", 100, 0, "")]
    [InlineData("none", 999, 0, "")]
    [InlineData("none", 1000, 0, "")]
    [InlineData("none", 9999, 0, "")]
    [InlineData("none", 10000, 0, "")]
    [InlineData("none", 32767, 0, "")]
    [InlineData("none", 32768, 0, "")]
    [InlineData("none", 65535, 0, "")]
    [InlineData("none", 999999, 0, "")]
    [InlineData("none", 2147483647, 0, "")]
    [InlineData("literal-control-none", -2147483648, 0, "literal")]
    [InlineData("literal-control-none", -1, 0, "literal")]
    [InlineData("literal-control-none", 0, 0, "literal")]
    [InlineData("literal-control-none", 1, 0, "literal")]
    [InlineData("literal-control-none", 9, 0, "literal")]
    [InlineData("literal-control-none", 10, 0, "literal")]
    [InlineData("literal-control-none", 99, 0, "literal")]
    [InlineData("literal-control-none", 100, 0, "literal")]
    [InlineData("literal-control-none", 999, 0, "literal")]
    [InlineData("literal-control-none", 1000, 0, "literal")]
    [InlineData("literal-control-none", 9999, 0, "literal")]
    [InlineData("literal-control-none", 10000, 0, "literal")]
    [InlineData("literal-control-none", 32767, 0, "literal")]
    [InlineData("literal-control-none", 32768, 0, "literal")]
    [InlineData("literal-control-none", 65535, 0, "literal")]
    [InlineData("literal-control-none", 999999, 0, "literal")]
    [InlineData("literal-control-none", 2147483647, 0, "literal")]
    [InlineData("literal-none", -2147483648, 0, "literal")]
    [InlineData("literal-none", -1, 0, "literal")]
    [InlineData("literal-none", 0, 0, "literal")]
    [InlineData("literal-none", 1, 0, "literal")]
    [InlineData("literal-none", 9, 0, "literal")]
    [InlineData("literal-none", 10, 0, "literal")]
    [InlineData("literal-none", 99, 0, "literal")]
    [InlineData("literal-none", 100, 0, "literal")]
    [InlineData("literal-none", 999, 0, "literal")]
    [InlineData("literal-none", 1000, 0, "literal")]
    [InlineData("literal-none", 9999, 0, "literal")]
    [InlineData("literal-none", 10000, 0, "literal")]
    [InlineData("literal-none", 32767, 0, "literal")]
    [InlineData("literal-none", 32768, 0, "literal")]
    [InlineData("literal-none", 65535, 0, "literal")]
    [InlineData("literal-none", 999999, 0, "literal")]
    [InlineData("literal-none", 2147483647, 0, "literal")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var literal = name.StartsWith("literal-", StringComparison.Ordinal);
        var abstractXml = (literal ? name.Substring(8) : name) switch
        {
            "control-none" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"none\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "none" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var abstractNum = new AbstractNum(abstractXml);
        if (literal) abstractNum.GetFirstChild<Level>()!.LevelText!.Val = "BEGINliteralEND";
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
