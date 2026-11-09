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

public sealed class CustomNumericListMarkerTests
{
    // Word-backed custom strings, MC branch selection and zero-override controls.
    [Theory]
    [InlineData("legal", -2147483648, "")]
    [InlineData("legal", -1, "-1")]
    [InlineData("legal", 0, "0")]
    [InlineData("legal", 1, "1")]
    [InlineData("legal", 9, "9")]
    [InlineData("legal", 10, "10")]
    [InlineData("legal", 99, "99")]
    [InlineData("legal", 100, "100")]
    [InlineData("legal", 999, "999")]
    [InlineData("legal", 1000, "1000")]
    [InlineData("legal", 9999, "9999")]
    [InlineData("legal", 10000, "10000")]
    [InlineData("legal", 32767, "32767")]
    [InlineData("legal", 32768, "32768")]
    [InlineData("legal", 65535, "65535")]
    [InlineData("legal", 999999, "999999")]
    [InlineData("legal", 2147483647, "2147483647")]
    [InlineData("nonzero", -2147483648, "")]
    [InlineData("nonzero", -1, "-1")]
    [InlineData("nonzero", 0, "0")]
    [InlineData("nonzero", 1, "1")]
    [InlineData("nonzero", 9, "9")]
    [InlineData("nonzero", 10, "10")]
    [InlineData("nonzero", 99, "99")]
    [InlineData("nonzero", 100, "100")]
    [InlineData("nonzero", 999, "999")]
    [InlineData("nonzero", 1000, "1000")]
    [InlineData("nonzero", 9999, "9999")]
    [InlineData("nonzero", 10000, "10000")]
    [InlineData("nonzero", 32767, "32767")]
    [InlineData("nonzero", 32768, "32768")]
    [InlineData("nonzero", 65535, "65535")]
    [InlineData("nonzero", 999999, "999999")]
    [InlineData("nonzero", 2147483647, "2147483647")]
    [InlineData("pad2", -2147483648, "")]
    [InlineData("pad2", -1, "")]
    [InlineData("pad2", 0, "00")]
    [InlineData("pad2", 1, "01")]
    [InlineData("pad2", 9, "09")]
    [InlineData("pad2", 10, "10")]
    [InlineData("pad2", 99, "99")]
    [InlineData("pad2", 100, "100")]
    [InlineData("pad2", 999, "999")]
    [InlineData("pad2", 1000, "1000")]
    [InlineData("pad2", 9999, "9999")]
    [InlineData("pad2", 10000, "10000")]
    [InlineData("pad2", 32767, "32767")]
    [InlineData("pad2", 32768, "32768")]
    [InlineData("pad2", 65535, "65535")]
    [InlineData("pad2", 999999, "999999")]
    [InlineData("pad2", 2147483647, "2147483647")]
    [InlineData("pad3", -2147483648, "")]
    [InlineData("pad3", -1, "")]
    [InlineData("pad3", 0, "000")]
    [InlineData("pad3", 1, "001")]
    [InlineData("pad3", 9, "009")]
    [InlineData("pad3", 10, "010")]
    [InlineData("pad3", 99, "099")]
    [InlineData("pad3", 100, "100")]
    [InlineData("pad3", 999, "999")]
    [InlineData("pad3", 1000, "1000")]
    [InlineData("pad3", 9999, "9999")]
    [InlineData("pad3", 10000, "10000")]
    [InlineData("pad3", 32767, "32767")]
    [InlineData("pad3", 32768, "32768")]
    [InlineData("pad3", 65535, "65535")]
    [InlineData("pad3", 999999, "999999")]
    [InlineData("pad3", 2147483647, "2147483647")]
    [InlineData("pad4", -2147483648, "")]
    [InlineData("pad4", -1, "")]
    [InlineData("pad4", 0, "0000")]
    [InlineData("pad4", 1, "0001")]
    [InlineData("pad4", 9, "0009")]
    [InlineData("pad4", 10, "0010")]
    [InlineData("pad4", 99, "0099")]
    [InlineData("pad4", 100, "0100")]
    [InlineData("pad4", 999, "0999")]
    [InlineData("pad4", 1000, "1000")]
    [InlineData("pad4", 9999, "9999")]
    [InlineData("pad4", 10000, "10000")]
    [InlineData("pad4", 32767, "32767")]
    [InlineData("pad4", 32768, "32768")]
    [InlineData("pad4", 65535, "65535")]
    [InlineData("pad4", 999999, "999999")]
    [InlineData("pad4", 2147483647, "2147483647")]
    [InlineData("pad5", -2147483648, "")]
    [InlineData("pad5", -1, "")]
    [InlineData("pad5", 0, "00000")]
    [InlineData("pad5", 1, "00001")]
    [InlineData("pad5", 9, "00009")]
    [InlineData("pad5", 10, "00010")]
    [InlineData("pad5", 99, "00099")]
    [InlineData("pad5", 100, "00100")]
    [InlineData("pad5", 999, "00999")]
    [InlineData("pad5", 1000, "01000")]
    [InlineData("pad5", 9999, "09999")]
    [InlineData("pad5", 10000, "10000")]
    [InlineData("pad5", 32767, "32767")]
    [InlineData("pad5", 32768, "32768")]
    [InlineData("pad5", 65535, "65535")]
    [InlineData("pad5", 999999, "999999")]
    [InlineData("pad5", 2147483647, "2147483647")]
    [InlineData("pad6", -2147483648, "")]
    [InlineData("pad6", -1, "-1")]
    [InlineData("pad6", 0, "0")]
    [InlineData("pad6", 1, "1")]
    [InlineData("pad6", 9, "9")]
    [InlineData("pad6", 10, "10")]
    [InlineData("pad6", 99, "99")]
    [InlineData("pad6", 100, "100")]
    [InlineData("pad6", 999, "999")]
    [InlineData("pad6", 1000, "1000")]
    [InlineData("pad6", 9999, "9999")]
    [InlineData("pad6", 10000, "10000")]
    [InlineData("pad6", 32767, "32767")]
    [InlineData("pad6", 32768, "32768")]
    [InlineData("pad6", 65535, "65535")]
    [InlineData("pad6", 999999, "999999")]
    [InlineData("pad6", 2147483647, "2147483647")]
    [InlineData("single", -2147483648, "")]
    [InlineData("single", -1, "-1")]
    [InlineData("single", 0, "0")]
    [InlineData("single", 1, "1")]
    [InlineData("single", 9, "9")]
    [InlineData("single", 10, "10")]
    [InlineData("single", 99, "99")]
    [InlineData("single", 100, "100")]
    [InlineData("single", 999, "999")]
    [InlineData("single", 1000, "1000")]
    [InlineData("single", 9999, "9999")]
    [InlineData("single", 10000, "10000")]
    [InlineData("single", 32767, "32767")]
    [InlineData("single", 32768, "32768")]
    [InlineData("single", 65535, "65535")]
    [InlineData("single", 999999, "999999")]
    [InlineData("single", 2147483647, "2147483647")]
    [InlineData("unknown", -2147483648, "")]
    [InlineData("unknown", -1, "-1")]
    [InlineData("unknown", 0, "0")]
    [InlineData("unknown", 1, "1")]
    [InlineData("unknown", 9, "9")]
    [InlineData("unknown", 10, "10")]
    [InlineData("unknown", 99, "99")]
    [InlineData("unknown", 100, "100")]
    [InlineData("unknown", 999, "999")]
    [InlineData("unknown", 1000, "1000")]
    [InlineData("unknown", 9999, "9999")]
    [InlineData("unknown", 10000, "10000")]
    [InlineData("unknown", 32767, "32767")]
    [InlineData("unknown", 32768, "32768")]
    [InlineData("unknown", 65535, "65535")]
    [InlineData("unknown", 999999, "999999")]
    [InlineData("unknown", 2147483647, "2147483647")]
    [InlineData("zeros", -2147483648, "")]
    [InlineData("zeros", -1, "-1")]
    [InlineData("zeros", 0, "0")]
    [InlineData("zeros", 1, "1")]
    [InlineData("zeros", 9, "9")]
    [InlineData("zeros", 10, "10")]
    [InlineData("zeros", 99, "99")]
    [InlineData("zeros", 100, "100")]
    [InlineData("zeros", 999, "999")]
    [InlineData("zeros", 1000, "1000")]
    [InlineData("zeros", 9999, "9999")]
    [InlineData("zeros", 10000, "10000")]
    [InlineData("zeros", 32767, "32767")]
    [InlineData("zeros", 32768, "32768")]
    [InlineData("zeros", 65535, "65535")]
    [InlineData("zeros", 999999, "999999")]
    [InlineData("zeros", 2147483647, "2147483647")]
    public void CustomFormatsMatchWord(string name, int start, string expected)
    {
        var levelXml = name switch
        {
            "legal" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"0001, 0002, 0003, ...\"/><w:isLgl/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "nonzero" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"1234, 1235, 1236, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "pad2" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"01, 02, 03, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "pad3" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"001, 002, 003, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "pad4" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"0001, 0002, 0003, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "pad5" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"00001, 00002, 00003, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "pad6" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"000001, 000002, 000003, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "single" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"0001\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "unknown" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"XYZ\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            "zeros" => "<w:lvl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"0000\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>",
            _ => throw new ArgumentException(name)
        };
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(new AbstractNum(new Level(levelXml)) { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 },
            new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = start }) { LevelIndex = 0 }) { NumberID = 2 });
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        Paragraph Item(int numId) => new(new ParagraphProperties(new NumberingProperties(
            new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numId })), new Run(new Text("Body")));
        main.Document = new Document(new Body(Item(1), Item(2)));
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            var matches = Regex.Matches(output, "BEGIN(.*?)END");
            Assert.NotEmpty(matches);
            Assert.Equal(expected, matches[matches.Count - 1].Groups[1].Value);
            Assert.Equal(original, numbering.OuterXml);
        }
    }
}
