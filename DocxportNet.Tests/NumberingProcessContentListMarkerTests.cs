using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace DocxportNet.Tests;

public sealed class NumberingProcessContentListMarkerTests
{
    [Theory]
    [InlineData("abstract", 7, "BEGIN")]
    [InlineData("instance", 7, "BEGIN")]
    [InlineData("level", 7, "BEGIN")]
    [InlineData("start", 7, "BEGIN")]
    [InlineData("label", 3, "CHOICE")]
    [InlineData("nested", 7, "BEGIN")]
    [InlineData("local-prefix", 7, "BEGIN")]
    [InlineData("wildcard", 7, "BEGIN")]
    [InlineData("ignored", 3, "BEGIN")]
    [InlineData("local-choice", 7, "BEGIN")]
    [InlineData("choice-wrapper", 7, "BEGIN")]
    public void IgnorableWrappersMatchSdkPreprocessingWithoutMutatingSource(string kind, int start, string label)
    {
        const string w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        string Wrap(string content) => "<x:wrap>" + content + "</x:wrap>";
        var level = "<w:lvl w:ilvl=\"0\"><w:start w:val=\"7\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"BEGIN%1END\"/></w:lvl>";
        if (kind is "start" or "nested" or "local-prefix" or "wildcard")
        {
            var property = Wrap("<w:start w:val=\"7\"/>");
            if (kind == "nested") property = Wrap(property);
            if (kind == "local-prefix") property = "<y:wrap xmlns:y=\"urn:marker-extension\"><w:start w:val=\"7\"/></y:wrap>";
            level = level.Replace("<w:start w:val=\"7\"/>", property);
        }
        if (kind == "label") level = level.Replace("w:val=\"7\"", "w:val=\"3\"").Replace("<w:lvlText w:val=\"BEGIN%1END\"/>", Wrap("<w:lvlText w:val=\"CHOICE%1END\"/>"));
        if (kind == "ignored") level = level.Replace("<w:start w:val=\"7\"/>", "<w:start w:val=\"3\"/>" + Wrap("<w:start w:val=\"7\"/>"));
        if (kind == "local-choice") level = level.Replace("<w:start w:val=\"7\"/>",
            "<x:wrap xmlns:q=\"" + w + "\"><mc:AlternateContent><mc:Choice Requires=\"q\"><w:start w:val=\"7\"/></mc:Choice><mc:Fallback><w:start w:val=\"2\"/></mc:Fallback></mc:AlternateContent></x:wrap>");
        if (kind == "choice-wrapper") level = level.Replace("<w:start w:val=\"7\"/>",
            "<mc:AlternateContent><mc:Choice Requires=\"w\">" + Wrap("<w:start w:val=\"7\"/>") + "</mc:Choice><mc:Fallback><w:start w:val=\"2\"/></mc:Fallback></mc:AlternateContent>");
        var abstractNum = "<w:abstractNum w:abstractNumId=\"1\">" + (kind == "level" ? Wrap(level) : level) + "</w:abstractNum>";
        var instance = "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num>";
        var process = kind == "ignored" ? "" : " mc:ProcessContent=\"x:" + (kind == "wildcard" ? "*" : "wrap") + "\"";
        var xml = "<w:numbering xmlns:w=\"" + w + "\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"urn:marker-extension\" mc:Ignorable=\"x\"" + process + ">" +
            (kind == "abstract" ? Wrap(abstractNum) : abstractNum) + (kind == "instance" ? Wrap(instance) : instance) + "</w:numbering>";
        using var stream = new MemoryStream();
        using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = doc.AddMainDocumentPart();
        var part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = new Numbering(xml);
        main.Document = new Document(new Body(Enumerable.Range(0, 3).Select(i => new Paragraph(
            new ParagraphProperties(new NumberingProperties(new NumberingLevelReference { Val = 0 }, new NumberingId { Val = 1 })),
            new Run(new Text("Item" + i))))));
        _ = part.Numbering.Descendants().ToArray();
        var original = part.Numbering.OuterXml;
        using var oracleStream = new MemoryStream();
        using (var copy = doc.Clone(oracleStream)) { }
        using (var zip = new ZipArchive(oracleStream, ZipArchiveMode.Update, true))
        {
            var path = part.Uri.OriginalString.TrimStart('/');
            zip.GetEntry(path)!.Delete();
            using var writer = new StreamWriter(zip.CreateEntry(path).Open());
            writer.Write(xml);
        }
        oracleStream.Position = 0;
        using var oracle = WordprocessingDocument.Open(oracleStream, false, new OpenSettings
        {
            MarkupCompatibilityProcessSettings = new(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
        });
        DxpITextVisitor[] Visitors() => [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()), new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        var expected = Enumerable.Range(start, 3).Select(n => label + n + "END").ToArray();
        foreach (var source in new[] { oracle, doc })
            foreach (var visitor in Visitors())
            {
                var output = DxpExport.ExportToString(source, visitor);
                if (visitor is DxpHtmlVisitor) output = new HtmlParser().ParseDocument(output).Body!.TextContent;
                Assert.Equal(expected, Regex.Matches(output, "(?:BEGIN|CHOICE)\\d+END").Select(m => m.Value));
                Assert.Equal(3, Regex.Matches(output, "Item[0-2]").Count);
            }
        Assert.Equal(original, part.Numbering.OuterXml);
    }
}
