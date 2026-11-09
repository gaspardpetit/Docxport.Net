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

public sealed class CustomDigitAliasListMarkerTests
{
    // Word captures for Hindi, Thai and Taiwanese custom digit aliases.
    [Theory]
    [InlineData("hindiNumbers", -2147483648, 0, "")]
    [InlineData("hindiNumbers", -1, 0, "")]
    [InlineData("hindiNumbers", 0, 0, "\u0966")]
    [InlineData("hindiNumbers", 1, 0, "\u0967")]
    [InlineData("hindiNumbers", 9, 0, "\u096f")]
    [InlineData("hindiNumbers", 10, 0, "\u0967\u0966")]
    [InlineData("hindiNumbers", 99, 0, "\u096f\u096f")]
    [InlineData("hindiNumbers", 100, 0, "\u0967\u0966\u0966")]
    [InlineData("hindiNumbers", 999, 0, "\u096f\u096f\u096f")]
    [InlineData("hindiNumbers", 1000, 0, "\u0967\u0966\u0966\u0966")]
    [InlineData("hindiNumbers", 9999, 0, "\u096f\u096f\u096f\u096f")]
    [InlineData("hindiNumbers", 10000, 0, "\u0967\u0966\u0966\u0966\u0966")]
    [InlineData("hindiNumbers", 32767, 0, "\u0969\u0968\u096d\u096c\u096d")]
    [InlineData("hindiNumbers", 32768, 0, "\u0967")]
    [InlineData("hindiNumbers", 65535, 0, "\u0967")]
    [InlineData("hindiNumbers", 999999, 0, "\u0967\u096c\u096f\u096e\u096f")]
    [InlineData("hindiNumbers", 2147483647, 0, "\u0967")]
    [InlineData("taiwaneseDigital", -2147483648, 0, "")]
    [InlineData("taiwaneseDigital", -1, 0, "")]
    [InlineData("taiwaneseDigital", 0, 0, "\u25cb")]
    [InlineData("taiwaneseDigital", 1, 0, "\u4e00")]
    [InlineData("taiwaneseDigital", 9, 0, "\u4e5d")]
    [InlineData("taiwaneseDigital", 10, 0, "\u4e00\u25cb")]
    [InlineData("taiwaneseDigital", 99, 0, "\u4e5d\u4e5d")]
    [InlineData("taiwaneseDigital", 100, 0, "\u4e00\u25cb\u25cb")]
    [InlineData("taiwaneseDigital", 999, 0, "\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseDigital", 1000, 0, "\u4e00\u25cb\u25cb\u25cb")]
    [InlineData("taiwaneseDigital", 9999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseDigital", 10000, 0, "\u4e00\u25cb\u25cb\u25cb\u25cb")]
    [InlineData("taiwaneseDigital", 32767, 0, "\u4e09\u4e8c\u4e03\u516d\u4e03")]
    [InlineData("taiwaneseDigital", 32768, 0, "\u4e09\u4e8c\u4e03\u516d\u516b")]
    [InlineData("taiwaneseDigital", 65535, 0, "\u516d\u4e94\u4e94\u4e09\u4e94")]
    [InlineData("taiwaneseDigital", 999999, 0, "\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d\u4e5d")]
    [InlineData("taiwaneseDigital", 2147483647, 0, "\u4e8c\u4e00\u56db\u4e03\u56db\u516b\u4e09\u516d\u56db\u4e03")]
    [InlineData("thaiNumbers", -2147483648, 0, "")]
    [InlineData("thaiNumbers", -1, 0, "")]
    [InlineData("thaiNumbers", 0, 0, "\u0e50")]
    [InlineData("thaiNumbers", 1, 0, "\u0e51")]
    [InlineData("thaiNumbers", 9, 0, "\u0e59")]
    [InlineData("thaiNumbers", 10, 0, "\u0e51\u0e50")]
    [InlineData("thaiNumbers", 99, 0, "\u0e59\u0e59")]
    [InlineData("thaiNumbers", 100, 0, "\u0e51\u0e50\u0e50")]
    [InlineData("thaiNumbers", 999, 0, "\u0e59\u0e59\u0e59")]
    [InlineData("thaiNumbers", 1000, 0, "\u0e51\u0e50\u0e50\u0e50")]
    [InlineData("thaiNumbers", 9999, 0, "\u0e59\u0e59\u0e59\u0e59")]
    [InlineData("thaiNumbers", 10000, 0, "\u0e51\u0e50\u0e50\u0e50\u0e50")]
    [InlineData("thaiNumbers", 32767, 0, "\u0e53\u0e52\u0e57\u0e56\u0e57")]
    [InlineData("thaiNumbers", 32768, 0, "\u0e51")]
    [InlineData("thaiNumbers", 65535, 0, "\u0e51")]
    [InlineData("thaiNumbers", 999999, 0, "\u0e51\u0e56\u0e59\u0e58\u0e59")]
    [InlineData("thaiNumbers", 2147483647, 0, "\u0e51")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var abstractXml = name switch
        {
            "hindiNumbers" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0967, \u0968, \u0969, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "taiwaneseDigital" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u4e00, \u4e00\u25cb, \u4e00\u25cb\u25cb(\u7e41), ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "thaiNumbers" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0e51, \u0e52, \u0e53, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(new AbstractNum(abstractXml),
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
