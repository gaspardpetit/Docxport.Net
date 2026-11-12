using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.Walker;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocLegalNumberingMarkerTests
{
    private static IEnumerable<object[]> OriginalWordDefinitions()
    {
        yield return new object[] { "bullet-legal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:isLgl/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTEND", "BEGIN/1END", "BEGIN/2END", "ROOTEND", "BEGIN/1END", "BEGIN/2END" } };
        yield return new object[] { "bullet-normal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTEND", "BEGIN/iEND", "BEGIN/iiEND", "ROOTEND", "BEGIN/iEND", "BEGIN/iiEND" } };
        yield return new object[] { "decimalZero-legal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"decimalZero\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:isLgl/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOT03END", "BEGIN03/1END", "BEGIN03/2END", "ROOT04END", "BEGIN04/1END", "BEGIN04/2END" } };
        yield return new object[] { "decimalZero-normal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"decimalZero\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOT03END", "BEGIN03/iEND", "BEGIN03/iiEND", "ROOT04END", "BEGIN04/iEND", "BEGIN04/iiEND" } };
        yield return new object[] { "none-legal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"none\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:isLgl/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTEND", "BEGIN/1END", "BEGIN/2END", "ROOTEND", "BEGIN/1END", "BEGIN/2END" } };
        yield return new object[] { "none-normal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"none\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTEND", "BEGIN/iEND", "BEGIN/iiEND", "ROOTEND", "BEGIN/iEND", "BEGIN/iiEND" } };
        yield return new object[] { "upperRoman-legal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:isLgl/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTIIIEND", "BEGIN3/1END", "BEGIN3/2END", "ROOTIVEND", "BEGIN4/1END", "BEGIN4/2END" } };
        yield return new object[] { "upperRoman-normal", "<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:abstractNum w:abstractNumId=\"1\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"3\"/><w:numFmt w:val=\"upperRoman\"/><w:lvlText w:val=\"ROOT%1END\"/></w:lvl><w:lvl w:ilvl=\"1\"><w:start w:val=\"1\"/><w:numFmt w:val=\"lowerRoman\"/><w:lvlText w:val=\"BEGIN%1/%2END\"/></w:lvl></w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num></w:numbering>", new int[] { 0, 1, 1, 0, 1, 1 }, new string[] { "ROOTIIIEND", "BEGINIII/iEND", "BEGINIII/iiEND", "ROOTIVEND", "BEGINIV/iEND", "BEGINIV/iiEND" } };
    }

    public static IEnumerable<object[]> WordDefinitions()
    {
        foreach (var row in OriginalWordDefinitions())
        {
            yield return row;
            if (((string)row[0]).EndsWith("-normal"))
                yield return new object[] { ((string)row[0]).Replace("-normal", "-disabled"),
                    ((string)row[1]).Replace("<w:lvlText w:val=\"BEGIN", "<w:isLgl w:val=\"false\"/><w:lvlText w:val=\"BEGIN"),
                    row[2], row[3] };
        }
    }

    [Theory]
    [MemberData(nameof(WordDefinitions))]
    public void LegalNumberingSurvivesEditableDocRoundTrip(string name, string xml, int[] levels, string[] expected)
    {
        using var stream = new MemoryStream();
        using var source = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = source.AddMainDocumentPart();
        var numbering = new Numbering(xml);
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
        main.Document = new Document(new Body(levels.Select((level, i) =>
            new Paragraph(new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = level }, new NumberingId { Val = 1 })),
                new Run(new Text($"Item{i}"))))));
        main.Document.Save();
        numbering.Save();
        source.Dispose();
        var binary = DxpDocExport.Export(stream.ToArray());
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
            Assert.Equal(name.EndsWith("-legal"), index.Lists.Definitions.Single().Levels.Single(x => x.Index == 1).IsLegal);
        var projectedBytes = DxpDocToDocx.Project(binary).DocxBytes;
        using var projectedStream = new MemoryStream(projectedBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var output = DxpExport.ExportToString(projected, visitor);
            if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
            Assert.Equal(expected, Regex.Matches(output, "(?:BEGIN|ROOT).*?END").Select(m => m.Value));
            Assert.Equal(levels.Length, Regex.Matches(output, @"Item\d+").Count);
        }
    }
}
