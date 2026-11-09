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

public sealed class CustomKanaAliasListMarkerTests
{
    // Word captures for Aiueo/Iroha half-width and full-width custom aliases.
    [Theory]
    [InlineData("aiueo", -2147483648, 0, "")]
    [InlineData("aiueo", -1, 0, "-1")]
    [InlineData("aiueo", 0, 0, "0")]
    [InlineData("aiueo", 1, 0, "\uff71")]
    [InlineData("aiueo", 9, 0, "\uff79")]
    [InlineData("aiueo", 10, 0, "\uff7a")]
    [InlineData("aiueo", 99, 0, "\uff77")]
    [InlineData("aiueo", 100, 0, "\uff78")]
    [InlineData("aiueo", 999, 0, "\uff91")]
    [InlineData("aiueo", 1000, 0, "\uff92")]
    [InlineData("aiueo", 9999, 0, "\uff81")]
    [InlineData("aiueo", 10000, 0, "\uff82")]
    [InlineData("aiueo", 32767, 0, "\uff7f")]
    [InlineData("aiueo", 32768, 0, "\uff80")]
    [InlineData("aiueo", 65535, 0, "\uff8f")]
    [InlineData("aiueo", 999999, 0, "\uff75")]
    [InlineData("aiueo", 2147483647, 0, "\uff75")]
    [InlineData("aiueoFullWidth", -2147483648, 0, "")]
    [InlineData("aiueoFullWidth", -1, 0, "-1")]
    [InlineData("aiueoFullWidth", 0, 0, "0")]
    [InlineData("aiueoFullWidth", 1, 0, "\u30a2")]
    [InlineData("aiueoFullWidth", 9, 0, "\u30b1")]
    [InlineData("aiueoFullWidth", 10, 0, "\u30b3")]
    [InlineData("aiueoFullWidth", 99, 0, "\u30ad")]
    [InlineData("aiueoFullWidth", 100, 0, "\u30af")]
    [InlineData("aiueoFullWidth", 999, 0, "\u30e0")]
    [InlineData("aiueoFullWidth", 1000, 0, "\u30e1")]
    [InlineData("aiueoFullWidth", 9999, 0, "\u30c1")]
    [InlineData("aiueoFullWidth", 10000, 0, "\u30c4")]
    [InlineData("aiueoFullWidth", 32767, 0, "\u30bd")]
    [InlineData("aiueoFullWidth", 32768, 0, "\u30bf")]
    [InlineData("aiueoFullWidth", 65535, 0, "\u30de")]
    [InlineData("aiueoFullWidth", 999999, 0, "\u30aa")]
    [InlineData("aiueoFullWidth", 2147483647, 0, "\u30aa")]
    [InlineData("iroha", -2147483648, 0, "")]
    [InlineData("iroha", -1, 0, "-1")]
    [InlineData("iroha", 0, 0, "0")]
    [InlineData("iroha", 1, 0, "\uff72")]
    [InlineData("iroha", 9, 0, "\uff98")]
    [InlineData("iroha", 10, 0, "\uff87")]
    [InlineData("iroha", 99, 0, "\uff8a")]
    [InlineData("iroha", 100, 0, "\uff86")]
    [InlineData("iroha", 999, 0, "\uff95")]
    [InlineData("iroha", 1000, 0, "\uff92")]
    [InlineData("iroha", 9999, 0, "\uff96")]
    [InlineData("iroha", 10000, 0, "\uff80")]
    [InlineData("iroha", 32767, 0, "\uff79")]
    [InlineData("iroha", 32768, 0, "\uff8c")]
    [InlineData("iroha", 65535, 0, "\uff96")]
    [InlineData("iroha", 999999, 0, "\uff96")]
    [InlineData("iroha", 2147483647, 0, "\uff79")]
    [InlineData("irohaFullWidth", -2147483648, 0, "")]
    [InlineData("irohaFullWidth", -1, 0, "-1")]
    [InlineData("irohaFullWidth", 0, 0, "0")]
    [InlineData("irohaFullWidth", 1, 0, "\u30a4")]
    [InlineData("irohaFullWidth", 9, 0, "\u30ea")]
    [InlineData("irohaFullWidth", 10, 0, "\u30cc")]
    [InlineData("irohaFullWidth", 99, 0, "\u30cf")]
    [InlineData("irohaFullWidth", 100, 0, "\u30cb")]
    [InlineData("irohaFullWidth", 999, 0, "\u30e6")]
    [InlineData("irohaFullWidth", 1000, 0, "\u30e1")]
    [InlineData("irohaFullWidth", 9999, 0, "\u30e8")]
    [InlineData("irohaFullWidth", 10000, 0, "\u30bf")]
    [InlineData("irohaFullWidth", 32767, 0, "\u30b1")]
    [InlineData("irohaFullWidth", 32768, 0, "\u30d5")]
    [InlineData("irohaFullWidth", 65535, 0, "\u30e8")]
    [InlineData("irohaFullWidth", 999999, 0, "\u30e8")]
    [InlineData("irohaFullWidth", 2147483647, 0, "\u30b1")]
    [InlineData("direct-min-aiueo", -2147483648, 0, "")]
    [InlineData("direct-min-aiueoFullWidth", -2147483648, 0, "")]
    [InlineData("direct-min-iroha", -2147483648, 0, "")]
    [InlineData("direct-min-irohaFullWidth", -2147483648, 0, "")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-min-", StringComparison.Ordinal) ? name.Substring(11) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "aiueo" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\uff71, \uff72, \uff73, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "aiueoFullWidth" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u30a2, \u30a4, \u30a6, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "iroha" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\uff72, \uff9b, \uff8a, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "irohaFullWidth" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u30a4, \u30ed, \u30cf, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
