using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocxportNet.Walker;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocBulletPlaceholderMarkerTests
{
    public static IEnumerable<object[]> WordDefinitions()
    {
        var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        var group = Path.Combine(root, "fixtures", "docx", "regressions", "list-markers", "BulletPlaceholders");
        using var manifest = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(group, "word-evidence.json")));
        foreach (var entry in manifest.RootElement.EnumerateObject())
        {
            using var doc = WordprocessingDocument.Open(Path.Combine(group, entry.Name + ".docx"), false);
            yield return new object[] { entry.Name, doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml,
                entry.Value.GetProperty("levels").EnumerateArray().Select(x => x.GetInt32()).ToArray(),
                entry.Value.GetProperty("labels").EnumerateArray().Select(x => x.GetString()!).ToArray() };
        }
        using var none = WordprocessingDocument.Open(Path.Combine(root, "fixtures", "docx", "regressions", "list-markers",
            "DocBulletPlaceholders", "source", "none-parent-none.docx"), false);
        yield return new object[] { "none-parent-none", none.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml,
            new[] { 0, 1, 1, 0, 1, 1 }, new[] { "ROOTEND", "BEGIN/END", "BEGIN/END", "ROOTEND", "BEGIN/END", "BEGIN/END" } };
    }

    [Theory]
    [InlineData((byte)0)]
    [InlineData((byte)0x17)]
    [InlineData((byte)0xff)]
    public void LevelReferencesUseBinaryPlaceholdersForEveryLabelFormat(byte format)
    {
        var lists = new DocListIndex([new DocListDefinition(1,
            [new DocListLevel(0, 1, format, "BEGIN%1/%2END", 0)])],
            [new DocListInstance(1, 1, new Dictionary<int, int>(), new Dictionary<int, DocListLevel>())]);
        var bytes = DocListWriter.Write(lists, new Dictionary<string, int>()).Levels;
        Assert.Equal(6, bytes[6]);
        Assert.Equal(8, bytes[7]);
        Assert.Equal("BEGIN\0/\u0001END", System.Text.Encoding.Unicode.GetString(bytes, 30, bytes.Length - 30));
    }

    [Theory]
    [MemberData(nameof(WordDefinitions))]
    public void BulletPlaceholdersSurviveEditableDocRoundTrip(string name, string xml, int[] levels, string[] expected)
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
        var projectedBytes = DxpDocToDocx.Project(binary).DocxBytes;
        using var projectedStream = new MemoryStream(projectedBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
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
