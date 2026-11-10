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

public sealed class CustomIdeographAliasListMarkerTests
{
    // Word captures for traditional, zodiac and enclosed ideograph custom aliases.
    [Theory]
    [InlineData("ideographEnclosedCircle", -2147483648, 0, "")]
    [InlineData("ideographEnclosedCircle", -1, 0, "-1")]
    [InlineData("ideographEnclosedCircle", 0, 0, "0")]
    [InlineData("ideographEnclosedCircle", 1, 0, "\u3220")]
    [InlineData("ideographEnclosedCircle", 9, 0, "\u3228")]
    [InlineData("ideographEnclosedCircle", 10, 0, "\u3229")]
    [InlineData("ideographEnclosedCircle", 99, 0, "99")]
    [InlineData("ideographEnclosedCircle", 100, 0, "100")]
    [InlineData("ideographEnclosedCircle", 999, 0, "999")]
    [InlineData("ideographEnclosedCircle", 1000, 0, "1000")]
    [InlineData("ideographEnclosedCircle", 9999, 0, "9999")]
    [InlineData("ideographEnclosedCircle", 10000, 0, "10000")]
    [InlineData("ideographEnclosedCircle", 32767, 0, "32767")]
    [InlineData("ideographEnclosedCircle", 32768, 0, "32768")]
    [InlineData("ideographEnclosedCircle", 65535, 0, "65535")]
    [InlineData("ideographEnclosedCircle", 999999, 0, "999999")]
    [InlineData("ideographEnclosedCircle", 2147483647, 0, "2147483647")]
    [InlineData("ideographTraditional", -2147483648, 0, "")]
    [InlineData("ideographTraditional", -1, 0, "-1")]
    [InlineData("ideographTraditional", 0, 0, "0")]
    [InlineData("ideographTraditional", 1, 0, "\u7532")]
    [InlineData("ideographTraditional", 9, 0, "\u58ec")]
    [InlineData("ideographTraditional", 10, 0, "\u7678")]
    [InlineData("ideographTraditional", 99, 0, "99")]
    [InlineData("ideographTraditional", 100, 0, "100")]
    [InlineData("ideographTraditional", 999, 0, "999")]
    [InlineData("ideographTraditional", 1000, 0, "1000")]
    [InlineData("ideographTraditional", 9999, 0, "9999")]
    [InlineData("ideographTraditional", 10000, 0, "10000")]
    [InlineData("ideographTraditional", 32767, 0, "32767")]
    [InlineData("ideographTraditional", 32768, 0, "32768")]
    [InlineData("ideographTraditional", 65535, 0, "65535")]
    [InlineData("ideographTraditional", 999999, 0, "999999")]
    [InlineData("ideographTraditional", 2147483647, 0, "2147483647")]
    [InlineData("ideographZodiac", -2147483648, 0, "")]
    [InlineData("ideographZodiac", -1, 0, "-1")]
    [InlineData("ideographZodiac", 0, 0, "0")]
    [InlineData("ideographZodiac", 1, 0, "\u5b50")]
    [InlineData("ideographZodiac", 9, 0, "\u7533")]
    [InlineData("ideographZodiac", 10, 0, "\u9149")]
    [InlineData("ideographZodiac", 99, 0, "99")]
    [InlineData("ideographZodiac", 100, 0, "100")]
    [InlineData("ideographZodiac", 999, 0, "999")]
    [InlineData("ideographZodiac", 1000, 0, "1000")]
    [InlineData("ideographZodiac", 9999, 0, "9999")]
    [InlineData("ideographZodiac", 10000, 0, "10000")]
    [InlineData("ideographZodiac", 32767, 0, "32767")]
    [InlineData("ideographZodiac", 32768, 0, "32768")]
    [InlineData("ideographZodiac", 65535, 0, "65535")]
    [InlineData("ideographZodiac", 999999, 0, "999999")]
    [InlineData("ideographZodiac", 2147483647, 0, "2147483647")]
    [InlineData("ideographZodiacTraditional", -2147483648, 0, "")]
    [InlineData("ideographZodiacTraditional", -1, 0, "-1")]
    [InlineData("ideographZodiacTraditional", 0, 0, "0")]
    [InlineData("ideographZodiacTraditional", 1, 0, "\u7532\u5b50")]
    [InlineData("ideographZodiacTraditional", 9, 0, "\u58ec\u7533")]
    [InlineData("ideographZodiacTraditional", 10, 0, "\u7678\u9149")]
    [InlineData("ideographZodiacTraditional", 99, 0, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 100, 0, "\u7678\u536f")]
    [InlineData("ideographZodiacTraditional", 999, 0, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 1000, 0, "\u7678\u536f")]
    [InlineData("ideographZodiacTraditional", 9999, 0, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 10000, 0, "\u7678\u536f")]
    [InlineData("ideographZodiacTraditional", 32767, 0, "\u5e9a\u5348")]
    [InlineData("ideographZodiacTraditional", 32768, 0, "\u8f9b\u672a")]
    [InlineData("ideographZodiacTraditional", 65535, 0, "\u620a\u5bc5")]
    [InlineData("ideographZodiacTraditional", 999999, 0, "\u58ec\u5bc5")]
    [InlineData("ideographZodiacTraditional", 2147483647, 0, "\u5e9a\u5348")]
    [InlineData("direct-ideographTraditional", -2147483648, 0, "")]
    [InlineData("direct-ideographTraditional", -1, 0, "-1")]
    [InlineData("direct-ideographTraditional", 0, 0, "0")]
    [InlineData("direct-ideographTraditional", 1, 0, "\u7532")]
    [InlineData("direct-ideographTraditional", 9, 0, "\u58ec")]
    [InlineData("direct-ideographTraditional", 10, 0, "\u7678")]
    [InlineData("direct-ideographTraditional", 99, 0, "99")]
    [InlineData("direct-ideographTraditional", 100, 0, "100")]
    [InlineData("direct-ideographTraditional", 999, 0, "999")]
    [InlineData("direct-ideographTraditional", 1000, 0, "1000")]
    [InlineData("direct-ideographTraditional", 9999, 0, "9999")]
    [InlineData("direct-ideographTraditional", 10000, 0, "10000")]
    [InlineData("direct-ideographTraditional", 32767, 0, "32767")]
    [InlineData("direct-ideographTraditional", 32768, 0, "32768")]
    [InlineData("direct-ideographTraditional", 65535, 0, "65535")]
    [InlineData("direct-ideographTraditional", 999999, 0, "999999")]
    [InlineData("direct-ideographTraditional", 2147483647, 0, "2147483647")]
    [InlineData("direct-ideographZodiac", -2147483648, 0, "")]
    [InlineData("direct-ideographZodiac", -1, 0, "-1")]
    [InlineData("direct-ideographZodiac", 0, 0, "0")]
    [InlineData("direct-ideographZodiac", 1, 0, "\u5b50")]
    [InlineData("direct-ideographZodiac", 9, 0, "\u7533")]
    [InlineData("direct-ideographZodiac", 10, 0, "\u9149")]
    [InlineData("direct-ideographZodiac", 99, 0, "99")]
    [InlineData("direct-ideographZodiac", 100, 0, "100")]
    [InlineData("direct-ideographZodiac", 999, 0, "999")]
    [InlineData("direct-ideographZodiac", 1000, 0, "1000")]
    [InlineData("direct-ideographZodiac", 9999, 0, "9999")]
    [InlineData("direct-ideographZodiac", 10000, 0, "10000")]
    [InlineData("direct-ideographZodiac", 32767, 0, "32767")]
    [InlineData("direct-ideographZodiac", 32768, 0, "32768")]
    [InlineData("direct-ideographZodiac", 65535, 0, "65535")]
    [InlineData("direct-ideographZodiac", 999999, 0, "999999")]
    [InlineData("direct-ideographZodiac", 2147483647, 0, "2147483647")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "ideographEnclosedCircle" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u3220, \u3221, \u3222, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ideographTraditional" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u7532, \u4e59, \u4e19, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ideographZodiac" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u5b50, \u4e11, \u5bc5, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ideographZodiacTraditional" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u7532\u5b50, \u4e59\u4e11, \u4e19\u5bc5, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
