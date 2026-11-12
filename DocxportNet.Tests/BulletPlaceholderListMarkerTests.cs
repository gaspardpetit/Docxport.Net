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

public sealed class BulletPlaceholderListMarkerTests
{
    public static IEnumerable<object[]> WordDefinitions()
    {
        yield return new object[] { "bullet-parent-bullet", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTEND", "BEGIN/END", "BEGIN/END", "ROOTEND", "BEGIN/END", "BEGIN/END" } };
        yield return new object[] { "bullet-parent-decimal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOT3END", "BEGIN3/END", "BEGIN3/END", "ROOT4END", "BEGIN4/END", "BEGIN4/END" } };
        yield return new object[] { "bullet-parent-upperRoman", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTIIIEND", "BEGINIII/END", "BEGINIII/END", "ROOTIVEND", "BEGINIV/END", "BEGINIV/END" } };
        yield return new object[] { "bullet-quotes", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"BEGINA&quot;B&#x27;CEND\"/><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 0, 0, 0, 0, 0 }, new string[] { "BEGINA\"B'CEND", "BEGINA\"B'CEND", "BEGINA\"B'CEND", "BEGINA\"B'CEND", "BEGINA\"B'CEND", "BEGINA\"B'CEND" } };
    }

    [Theory]
    [MemberData(nameof(WordDefinitions))]
    public void BulletPlaceholdersAndLiteralControlMatchWord(string name, string xml, int[] levels, string[] expected)
    {
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var numbering = new Numbering(xml);
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        main.Document = new Document(new Body(levels.Select((level, i) =>
            new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = level }, new NumberingId { Val = 1 })),
                new Run(new Text($"Item{i}"))))));
        var original = numbering.OuterXml;
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            Assert.Equal(expected, Regex.Matches(output, "(?:BEGIN|ROOT).*?END").Select(m => m.Value));
            Assert.Equal(levels.Length, Regex.Matches(output, @"Item\d+").Count);
            // Raw-XML fixtures are lazily parsed by the SDK, which can change empty-tag spacing.
            Assert.True(System.Xml.Linq.XNode.DeepEquals(System.Xml.Linq.XElement.Parse(original),
                System.Xml.Linq.XElement.Parse(numbering.OuterXml)), "Source numbering XML must remain equivalent.");
        }
    }
}
