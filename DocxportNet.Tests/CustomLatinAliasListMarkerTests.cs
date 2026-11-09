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

public sealed class CustomLatinAliasListMarkerTests
{
    // Word captures for the four Latin letter and Roman custom aliases.
    [Theory]
    [InlineData("lowerLetter", -2147483648, 0, "")]
    [InlineData("lowerLetter", -1, 0, "")]
    [InlineData("lowerLetter", 0, 0, " ")]
    [InlineData("lowerLetter", 1, 0, "a")]
    [InlineData("lowerLetter", 9, 0, "i")]
    [InlineData("lowerLetter", 10, 0, "j")]
    [InlineData("lowerLetter", 99, 0, "uuuu")]
    [InlineData("lowerLetter", 100, 0, "vvvv")]
    [InlineData("lowerLetter", 999, 0, "kkkkkkkkk")]
    [InlineData("lowerLetter", 1000, 0, "lllllllll")]
    [InlineData("lowerLetter", 9999, 0, "ooooooooooooooooooooooooo")]
    [InlineData("lowerLetter", 10000, 0, "ppppppppppppppppppppppppp")]
    [InlineData("lowerLetter", 32767, 0, "g")]
    [InlineData("lowerLetter", 32768, 0, "h")]
    [InlineData("lowerLetter", 65535, 0, "o")]
    [InlineData("lowerLetter", 999999, 0, "mm")]
    [InlineData("lowerLetter", 2147483647, 0, "wwwww")]
    [InlineData("lowerRoman", -2147483648, 0, "")]
    [InlineData("lowerRoman", -1, 0, "")]
    [InlineData("lowerRoman", 0, 0, " ")]
    [InlineData("lowerRoman", 1, 0, "i")]
    [InlineData("lowerRoman", 9, 0, "ix")]
    [InlineData("lowerRoman", 10, 0, "x")]
    [InlineData("lowerRoman", 99, 0, "xcix")]
    [InlineData("lowerRoman", 100, 0, "c")]
    [InlineData("lowerRoman", 999, 0, "cmxcix")]
    [InlineData("lowerRoman", 1000, 0, "m")]
    [InlineData("lowerRoman", 9999, 0, "mmmmmmmmmcmxcix")]
    [InlineData("lowerRoman", 10000, 0, "mmmmmmmmmm")]
    [InlineData("lowerRoman", 32767, 0, "mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmdcclxvii")]
    [InlineData("lowerRoman", 32768, 0, "i")]
    [InlineData("lowerRoman", 65535, 0, "i")]
    [InlineData("lowerRoman", 999999, 0, "mmmmmmmmmmmmmmmmcmlxxxix")]
    [InlineData("lowerRoman", 2147483647, 0, "i")]
    [InlineData("upperLetter", -2147483648, 0, "")]
    [InlineData("upperLetter", -1, 0, "")]
    [InlineData("upperLetter", 0, 0, " ")]
    [InlineData("upperLetter", 1, 0, "A")]
    [InlineData("upperLetter", 9, 0, "I")]
    [InlineData("upperLetter", 10, 0, "J")]
    [InlineData("upperLetter", 99, 0, "UUUU")]
    [InlineData("upperLetter", 100, 0, "VVVV")]
    [InlineData("upperLetter", 999, 0, "KKKKKKKKK")]
    [InlineData("upperLetter", 1000, 0, "LLLLLLLLL")]
    [InlineData("upperLetter", 9999, 0, "OOOOOOOOOOOOOOOOOOOOOOOOO")]
    [InlineData("upperLetter", 10000, 0, "PPPPPPPPPPPPPPPPPPPPPPPPP")]
    [InlineData("upperLetter", 32767, 0, "G")]
    [InlineData("upperLetter", 32768, 0, "H")]
    [InlineData("upperLetter", 65535, 0, "O")]
    [InlineData("upperLetter", 999999, 0, "MM")]
    [InlineData("upperLetter", 2147483647, 0, "WWWWW")]
    [InlineData("upperRoman", -2147483648, 0, "")]
    [InlineData("upperRoman", -1, 0, "")]
    [InlineData("upperRoman", 0, 0, " ")]
    [InlineData("upperRoman", 1, 0, "I")]
    [InlineData("upperRoman", 9, 0, "IX")]
    [InlineData("upperRoman", 10, 0, "X")]
    [InlineData("upperRoman", 99, 0, "XCIX")]
    [InlineData("upperRoman", 100, 0, "C")]
    [InlineData("upperRoman", 999, 0, "CMXCIX")]
    [InlineData("upperRoman", 1000, 0, "M")]
    [InlineData("upperRoman", 9999, 0, "MMMMMMMMMCMXCIX")]
    [InlineData("upperRoman", 10000, 0, "MMMMMMMMMM")]
    [InlineData("upperRoman", 32767, 0, "MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMDCCLXVII")]
    [InlineData("upperRoman", 32768, 0, "I")]
    [InlineData("upperRoman", 65535, 0, "I")]
    [InlineData("upperRoman", 999999, 0, "MMMMMMMMMMMMMMMMCMLXXXIX")]
    [InlineData("upperRoman", 2147483647, 0, "I")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var abstractXml = name switch
        {
            "lowerLetter" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"a\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "lowerRoman" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"i\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "upperLetter" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"A\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "upperRoman" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"I\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
