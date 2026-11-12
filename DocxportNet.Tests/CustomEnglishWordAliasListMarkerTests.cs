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

public sealed class CustomEnglishWordAliasListMarkerTests
{
    // English word aliases and direct formats, checked against Word captures.
    [Theory]
    [InlineData("cardinalText", -2147483648, 0, "")]
    [InlineData("cardinalText", -1, 0, "")]
    [InlineData("cardinalText", 0, 0, "Zero")]
    [InlineData("cardinalText", 1, 0, "One")]
    [InlineData("cardinalText", 9, 0, "Nine")]
    [InlineData("cardinalText", 10, 0, "Ten")]
    [InlineData("cardinalText", 99, 0, "Ninety-nine")]
    [InlineData("cardinalText", 100, 0, "One hundred")]
    [InlineData("cardinalText", 999, 0, "Nine hundred ninety-nine")]
    [InlineData("cardinalText", 1000, 0, "One thousand")]
    [InlineData("cardinalText", 9999, 0, "Nine thousand nine hundred ninety-nine")]
    [InlineData("cardinalText", 10000, 0, "Ten thousand")]
    [InlineData("cardinalText", 32767, 0, "Thirty-two thousand seven hundred sixty-seven")]
    [InlineData("cardinalText", 32768, 0, "Thirty-two thousand seven hundred sixty-eight")]
    [InlineData("cardinalText", 65535, 0, "Sixty-five thousand five hundred thirty-five")]
    [InlineData("cardinalText", 999999, 0, "Nine hundred ninety-nine thousand nine hundred ninety-nine")]
    [InlineData("cardinalText", 2147483647, 0, "")]
    [InlineData("ordinalText", -2147483648, 0, "")]
    [InlineData("ordinalText", -1, 0, "")]
    [InlineData("ordinalText", 0, 0, "Zeroth")]
    [InlineData("ordinalText", 1, 0, "First")]
    [InlineData("ordinalText", 9, 0, "Ninth")]
    [InlineData("ordinalText", 10, 0, "Tenth")]
    [InlineData("ordinalText", 99, 0, "Ninety-ninth")]
    [InlineData("ordinalText", 100, 0, "One hundredth")]
    [InlineData("ordinalText", 999, 0, "Nine hundred ninety-ninth")]
    [InlineData("ordinalText", 1000, 0, "One thousandth")]
    [InlineData("ordinalText", 9999, 0, "Nine thousand nine hundred ninety-ninth")]
    [InlineData("ordinalText", 10000, 0, "Ten thousandth")]
    [InlineData("ordinalText", 32767, 0, "Thirty-two thousand seven hundred sixty-seventh")]
    [InlineData("ordinalText", 32768, 0, "Thirty-two thousand seven hundred sixty-eighth")]
    [InlineData("ordinalText", 65535, 0, "Sixty-five thousand five hundred thirty-fifth")]
    [InlineData("ordinalText", 999999, 0, "Nine hundred ninety-nine thousand nine hundred ninety-ninth")]
    [InlineData("ordinalText", 2147483647, 0, "")]
    [InlineData("direct-cardinalText", -2147483648, 0, "")]
    [InlineData("direct-cardinalText", -1, 0, "")]
    [InlineData("direct-cardinalText", 0, 0, "Zero")]
    [InlineData("direct-cardinalText", 1, 0, "One")]
    [InlineData("direct-cardinalText", 9, 0, "Nine")]
    [InlineData("direct-cardinalText", 10, 0, "Ten")]
    [InlineData("direct-cardinalText", 99, 0, "Ninety-nine")]
    [InlineData("direct-cardinalText", 100, 0, "One hundred")]
    [InlineData("direct-cardinalText", 999, 0, "Nine hundred ninety-nine")]
    [InlineData("direct-cardinalText", 1000, 0, "One thousand")]
    [InlineData("direct-cardinalText", 9999, 0, "Nine thousand nine hundred ninety-nine")]
    [InlineData("direct-cardinalText", 10000, 0, "Ten thousand")]
    [InlineData("direct-cardinalText", 32767, 0, "Thirty-two thousand seven hundred sixty-seven")]
    [InlineData("direct-cardinalText", 32768, 0, "Thirty-two thousand seven hundred sixty-eight")]
    [InlineData("direct-cardinalText", 65535, 0, "Sixty-five thousand five hundred thirty-five")]
    [InlineData("direct-cardinalText", 999999, 0, "Nine hundred ninety-nine thousand nine hundred ninety-nine")]
    [InlineData("direct-cardinalText", 2147483647, 0, "")]
    [InlineData("direct-ordinalText", -2147483648, 0, "")]
    [InlineData("direct-ordinalText", -1, 0, "")]
    [InlineData("direct-ordinalText", 0, 0, "Zeroth")]
    [InlineData("direct-ordinalText", 1, 0, "First")]
    [InlineData("direct-ordinalText", 9, 0, "Ninth")]
    [InlineData("direct-ordinalText", 10, 0, "Tenth")]
    [InlineData("direct-ordinalText", 99, 0, "Ninety-ninth")]
    [InlineData("direct-ordinalText", 100, 0, "One hundredth")]
    [InlineData("direct-ordinalText", 999, 0, "Nine hundred ninety-ninth")]
    [InlineData("direct-ordinalText", 1000, 0, "One thousandth")]
    [InlineData("direct-ordinalText", 9999, 0, "Nine thousand nine hundred ninety-ninth")]
    [InlineData("direct-ordinalText", 10000, 0, "Ten thousandth")]
    [InlineData("direct-ordinalText", 32767, 0, "Thirty-two thousand seven hundred sixty-seventh")]
    [InlineData("direct-ordinalText", 32768, 0, "Thirty-two thousand seven hundred sixty-eighth")]
    [InlineData("direct-ordinalText", 65535, 0, "Sixty-five thousand five hundred thirty-fifth")]
    [InlineData("direct-ordinalText", 999999, 0, "Nine hundred ninety-nine thousand nine hundred ninety-ninth")]
    [InlineData("direct-ordinalText", 2147483647, 0, "")]
    [InlineData("boundary-cardinalText", -2147483648, 0, "")]
    [InlineData("boundary-cardinalText", -2, 0, "")]
    [InlineData("boundary-cardinalText", -1, 0, "")]
    [InlineData("boundary-cardinalText", 0, 0, "Zero")]
    [InlineData("boundary-cardinalText", 32767, 0, "Thirty-two thousand seven hundred sixty-seven")]
    [InlineData("boundary-cardinalText", 32768, 0, "Thirty-two thousand seven hundred sixty-eight")]
    [InlineData("boundary-cardinalText", 32769, 0, "Thirty-two thousand seven hundred sixty-nine")]
    [InlineData("boundary-cardinalText", 33000, 0, "Thirty-three thousand")]
    [InlineData("boundary-cardinalText", 33001, 0, "Thirty-three thousand one")]
    [InlineData("boundary-cardinalText", 100000, 0, "One hundred thousand")]
    [InlineData("boundary-cardinalText", 100001, 0, "One hundred thousand one")]
    [InlineData("boundary-cardinalText", 999000, 0, "Nine hundred ninety-nine thousand")]
    [InlineData("boundary-cardinalText", 999001, 0, "Nine hundred ninety-nine thousand one")]
    [InlineData("boundary-cardinalText", 999999, 0, "Nine hundred ninety-nine thousand nine hundred ninety-nine")]
    [InlineData("boundary-cardinalText", 1000000, 0, "")]
    [InlineData("boundary-cardinalText", 1000001, 0, "")]
    [InlineData("boundary-cardinalText", 2147483647, 0, "")]
    [InlineData("boundary-ordinalText", -2147483648, 0, "")]
    [InlineData("boundary-ordinalText", -2, 0, "")]
    [InlineData("boundary-ordinalText", -1, 0, "")]
    [InlineData("boundary-ordinalText", 0, 0, "Zeroth")]
    [InlineData("boundary-ordinalText", 32767, 0, "Thirty-two thousand seven hundred sixty-seventh")]
    [InlineData("boundary-ordinalText", 32768, 0, "Thirty-two thousand seven hundred sixty-eighth")]
    [InlineData("boundary-ordinalText", 32769, 0, "Thirty-two thousand seven hundred sixty-ninth")]
    [InlineData("boundary-ordinalText", 33000, 0, "Thirty-three thousandth")]
    [InlineData("boundary-ordinalText", 33001, 0, "Thirty-three thousand first")]
    [InlineData("boundary-ordinalText", 100000, 0, "One hundred thousandth")]
    [InlineData("boundary-ordinalText", 100001, 0, "One hundred thousand first")]
    [InlineData("boundary-ordinalText", 999000, 0, "Nine hundred ninety-nine thousandth")]
    [InlineData("boundary-ordinalText", 999001, 0, "Nine hundred ninety-nine thousand first")]
    [InlineData("boundary-ordinalText", 999999, 0, "Nine hundred ninety-nine thousand nine hundred ninety-ninth")]
    [InlineData("boundary-ordinalText", 1000000, 0, "")]
    [InlineData("boundary-ordinalText", 1000001, 0, "")]
    [InlineData("boundary-ordinalText", 2147483647, 0, "")]
    public void AliasesMatchWord(string name, int start, int level, string? expected)
    {
        var directFormat = name.StartsWith("direct-", StringComparison.Ordinal) ? name.Substring(7)
            : name.StartsWith("boundary-", StringComparison.Ordinal) ? name.Substring(9) : null;
        var abstractXml = (directFormat ?? name) switch
        {
            "cardinalText" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"One, Two, Three, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
            "ordinalText" => "<w:abstractNum xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:x=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:fake=\"urn:unrecognized\" w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"custom\" w:format=\"First, Second, Third, ...\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl></w:abstractNum>",
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
            new NumberingLevelReference { Val = level }, new NumberingId { Val = 2 })), new Run(new RunProperties(new Languages { Val = "en-US" }), new Text("Body")));
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

    [Theory]
    [InlineData(false, 32768, "32768")]
    [InlineData(true, 32768, "32768")]
    [InlineData(false, 999999, "999999")]
    [InlineData(true, 1000000, "1000000")]
    [InlineData(false, -1, "minus one")]
    [InlineData(true, -1, "minus first")]
    public void SharedEnglishFieldFormattingRetainsItsRange(bool ordinal, int value, string expected)
    {
        Assert.Equal(expected, DxpListMarkerFormatter.FormatNumber(value,
            ordinal ? NumberFormatValues.OrdinalText : NumberFormatValues.CardinalText,
            System.Globalization.CultureInfo.GetCultureInfo("en-US")));
    }

}
