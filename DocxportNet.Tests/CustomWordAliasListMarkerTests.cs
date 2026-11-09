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

public sealed class CustomWordAliasListMarkerTests
{
    // Word captures for counting-word aliases and noncanonical whitespace controls.
    [Theory]
    [InlineData("hindiCounting", -2147483648, 0, "")]
    [InlineData("hindiCounting", -1, 0, "")]
    [InlineData("hindiCounting", 0, 0, "\u0936\u0942\u0928\u094d\u092f")]
    [InlineData("hindiCounting", 1, 0, "\u090f\u0915")]
    [InlineData("hindiCounting", 9, 0, "\u0928\u094c")]
    [InlineData("hindiCounting", 10, 0, "\u0926\u0938")]
    [InlineData("hindiCounting", 99, 0, "\u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947")]
    [InlineData("hindiCounting", 100, 0, "\u090f\u0915 \u0938\u094c")]
    [InlineData("hindiCounting", 999, 0, "\u0928\u094c \u0938\u094c \u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947")]
    [InlineData("hindiCounting", 1000, 0, "\u090f\u0915 \u0939\u091c\u093e\u0930")]
    [InlineData("hindiCounting", 9999, 0, "\u0928\u094c \u0939\u091c\u093e\u0930 \u0928\u094c \u0938\u094c \u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947")]
    [InlineData("hindiCounting", 10000, 0, "\u090f\u0915")]
    [InlineData("hindiCounting", 32767, 0, "\u0926\u094b \u0939\u091c\u093e\u0930 \u0938\u093e\u0924 \u0938\u094c \u0938\u0924\u094d\u0924\u0930")]
    [InlineData("hindiCounting", 32768, 0, "\u0926\u094b \u0939\u091c\u093e\u0930 \u0938\u093e\u0924 \u0938\u094c \u090f\u0915\u0939\u0924\u094d\u0924\u0930")]
    [InlineData("hindiCounting", 65535, 0, "\u092a\u093e\u0901\u091a \u0939\u091c\u093e\u0930 \u092a\u093e\u0901\u091a \u0938\u094c \u0907\u0915\u094d\u0924\u093e\u0932\u0940\u0938")]
    [InlineData("hindiCounting", 999999, 0, "\u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947")]
    [InlineData("hindiCounting", 2147483647, 0, "\u0906\u0920 \u0939\u091c\u093e\u0930 \u091a\u093e\u0930 \u0938\u094c \u0938\u094b\u0932\u0939")]
    [InlineData("thaiCounting", -2147483648, 0, "")]
    [InlineData("thaiCounting", -1, 0, "\u0e25\u0e1a\u0e2b\u0e19\u0e36\u0e48\u0e07")]
    [InlineData("thaiCounting", 0, 0, "\u0e28\u0e39\u0e19\u0e22\u0e4c")]
    [InlineData("thaiCounting", 1, 0, "\u0e2b\u0e19\u0e36\u0e48\u0e07")]
    [InlineData("thaiCounting", 9, 0, "\u0e40\u0e01\u0e49\u0e32")]
    [InlineData("thaiCounting", 10, 0, "\u0e2a\u0e34\u0e1a")]
    [InlineData("thaiCounting", 99, 0, "\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e40\u0e01\u0e49\u0e32")]
    [InlineData("thaiCounting", 100, 0, "\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e23\u0e49\u0e2d\u0e22")]
    [InlineData("thaiCounting", 999, 0, "\u0e40\u0e01\u0e49\u0e32\u0e23\u0e49\u0e2d\u0e22\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e40\u0e01\u0e49\u0e32")]
    [InlineData("thaiCounting", 1000, 0, "\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e1e\u0e31\u0e19")]
    [InlineData("thaiCounting", 9999, 0, "\u0e40\u0e01\u0e49\u0e32\u0e1e\u0e31\u0e19\u0e40\u0e01\u0e49\u0e32\u0e23\u0e49\u0e2d\u0e22\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e40\u0e01\u0e49\u0e32")]
    [InlineData("thaiCounting", 10000, 0, "\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e2b\u0e21\u0e37\u0e48\u0e19")]
    [InlineData("thaiCounting", 32767, 0, "\u0e2a\u0e32\u0e21\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e2a\u0e2d\u0e07\u0e1e\u0e31\u0e19\u0e40\u0e08\u0e47\u0e14\u0e23\u0e49\u0e2d\u0e22\u0e2b\u0e01\u0e2a\u0e34\u0e1a\u0e40\u0e08\u0e47\u0e14")]
    [InlineData("thaiCounting", 32768, 0, "\u0e2a\u0e32\u0e21\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e2a\u0e2d\u0e07\u0e1e\u0e31\u0e19\u0e40\u0e08\u0e47\u0e14\u0e23\u0e49\u0e2d\u0e22\u0e2b\u0e01\u0e2a\u0e34\u0e1a\u0e41\u0e1b\u0e14")]
    [InlineData("thaiCounting", 65535, 0, "\u0e2b\u0e01\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e2b\u0e49\u0e32\u0e1e\u0e31\u0e19\u0e2b\u0e49\u0e32\u0e23\u0e49\u0e2d\u0e22\u0e2a\u0e32\u0e21\u0e2a\u0e34\u0e1a\u0e2b\u0e49\u0e32")]
    [InlineData("thaiCounting", 999999, 0, "\u0e40\u0e01\u0e49\u0e32\u0e41\u0e2a\u0e19\u0e40\u0e01\u0e49\u0e32\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e40\u0e01\u0e49\u0e32\u0e1e\u0e31\u0e19\u0e40\u0e01\u0e49\u0e32\u0e23\u0e49\u0e2d\u0e22\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e40\u0e01\u0e49\u0e32")]
    [InlineData("thaiCounting", 2147483647, 0, "\u0e28\u0e39\u0e19\u0e22\u0e4c")]
    [InlineData("thaiCounting-whitespace", -2147483648, 0, "")]
    [InlineData("thaiCounting-whitespace", -1, 0, "-1")]
    [InlineData("thaiCounting-whitespace", 0, 0, "0")]
    [InlineData("thaiCounting-whitespace", 1, 0, "1")]
    [InlineData("thaiCounting-whitespace", 9, 0, "9")]
    [InlineData("thaiCounting-whitespace", 10, 0, "10")]
    [InlineData("thaiCounting-whitespace", 99, 0, "99")]
    [InlineData("thaiCounting-whitespace", 100, 0, "100")]
    [InlineData("thaiCounting-whitespace", 999, 0, "999")]
    [InlineData("thaiCounting-whitespace", 1000, 0, "1000")]
    [InlineData("thaiCounting-whitespace", 9999, 0, "9999")]
    [InlineData("thaiCounting-whitespace", 10000, 0, "10000")]
    [InlineData("thaiCounting-whitespace", 32767, 0, "32767")]
    [InlineData("thaiCounting-whitespace", 32768, 0, "32768")]
    [InlineData("thaiCounting-whitespace", 65535, 0, "65535")]
    [InlineData("thaiCounting-whitespace", 999999, 0, "999999")]
    [InlineData("thaiCounting-whitespace", 2147483647, 0, "2147483647")]
    [InlineData("vietnameseCounting", -2147483648, 0, "")]
    [InlineData("vietnameseCounting", -1, 0, "")]
    [InlineData("vietnameseCounting", 0, 0, "kh\u00f4ng")]
    [InlineData("vietnameseCounting", 1, 0, "m\u00f4\u0323t")]
    [InlineData("vietnameseCounting", 9, 0, "chi\u0301n")]
    [InlineData("vietnameseCounting", 10, 0, "m\u01b0\u01a1\u0300i")]
    [InlineData("vietnameseCounting", 99, 0, "chi\u0301n m\u01b0\u01a1i chi\u0301n")]
    [InlineData("vietnameseCounting", 100, 0, "m\u00f4\u0323t tr\u0103m")]
    [InlineData("vietnameseCounting", 999, 0, "chi\u0301n tr\u0103m chi\u0301n m\u01b0\u01a1i chi\u0301n")]
    [InlineData("vietnameseCounting", 1000, 0, "m\u00f4\u0323t nga\u0300n")]
    [InlineData("vietnameseCounting", 9999, 0, "chi\u0301n tr\u0103m chi\u0301n m\u01b0\u01a1i chi\u0301n")]
    [InlineData("vietnameseCounting", 10000, 0, "m\u00f4\u0323t nga\u0300n")]
    [InlineData("vietnameseCounting", 32767, 0, "ba\u0309y tr\u0103m sa\u0301u m\u01b0\u01a1i ba\u0309y")]
    [InlineData("vietnameseCounting", 32768, 0, "ba\u0309y tr\u0103m sa\u0301u m\u01b0\u01a1i ta\u0301m")]
    [InlineData("vietnameseCounting", 65535, 0, "n\u0103m tr\u0103m ba m\u01b0\u01a1i l\u0103m")]
    [InlineData("vietnameseCounting", 999999, 0, "chi\u0301n tr\u0103m chi\u0301n m\u01b0\u01a1i chi\u0301n")]
    [InlineData("vietnameseCounting", 2147483647, 0, "sa\u0301u tr\u0103m b\u00f4\u0301n m\u01b0\u01a1i ba\u0309y")]
    [InlineData("vietnameseCounting-whitespace", -2147483648, 0, "")]
    [InlineData("vietnameseCounting-whitespace", -1, 0, "-1")]
    [InlineData("vietnameseCounting-whitespace", 0, 0, "0")]
    [InlineData("vietnameseCounting-whitespace", 1, 0, "1")]
    [InlineData("vietnameseCounting-whitespace", 9, 0, "9")]
    [InlineData("vietnameseCounting-whitespace", 10, 0, "10")]
    [InlineData("vietnameseCounting-whitespace", 99, 0, "99")]
    [InlineData("vietnameseCounting-whitespace", 100, 0, "100")]
    [InlineData("vietnameseCounting-whitespace", 999, 0, "999")]
    [InlineData("vietnameseCounting-whitespace", 1000, 0, "1000")]
    [InlineData("vietnameseCounting-whitespace", 9999, 0, "9999")]
    [InlineData("vietnameseCounting-whitespace", 10000, 0, "10000")]
    [InlineData("vietnameseCounting-whitespace", 32767, 0, "32767")]
    [InlineData("vietnameseCounting-whitespace", 32768, 0, "32768")]
    [InlineData("vietnameseCounting-whitespace", 65535, 0, "65535")]
    [InlineData("vietnameseCounting-whitespace", 999999, 0, "999999")]
    [InlineData("vietnameseCounting-whitespace", 2147483647, 0, "2147483647")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var abstractXml = name switch
        {
            "hindiCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u090f\u0915, \u0926\u094b, \u0924\u0940\u0928, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "thaiCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0e2b\u0e19\u0e36\u0e48\u0e07, \u0e2a\u0e2d\u0e07, \u0e2a\u0e32\u0e21, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "thaiCounting-whitespace" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"\u0e2b\u0e19\u0e36\u0e48\u0e07, \u0e2a\u0e2d\u0e07,\n  \u0e2a\u0e32\u0e21, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "vietnameseCounting" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"m\u00f4\u0323t, hai, ba, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "vietnameseCounting-whitespace" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"m\u00f4\u0323t, hai,\n  ba, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
