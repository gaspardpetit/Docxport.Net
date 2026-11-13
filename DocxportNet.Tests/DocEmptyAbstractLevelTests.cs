using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocEmptyAbstractLevelTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (var mode in new[] { "none", "start", "format", "both" })
                yield return new object[] { depth, mode, false };
        foreach (var mode in new[] { "none", "start", "format", "both" })
            yield return new object[] { 0, mode, true };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void EmptyDefinitionsPreserveBlankCountersAndReplacementActivation(int depth, string mode, bool root)
    {
        var bytes = CreateSource(depth, mode, root);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var levels = Assert.Single(index.Lists.Definitions).Levels;
            Assert.Equal(root ? 1 : 9, levels.Count);
            foreach (var level in levels)
            {
                Assert.Equal(root ? 2 : 0, level.StartAt);
                Assert.Equal(root ? "BEGIN%1END" : "", level.NumberText);
            }
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var expected = Labels(depth, mode, root);
        foreach (var document in new[] { source, projected })
        {
            DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
            foreach (var visitor in visitors)
            {
                var output = DxpExport.ExportToString(document, visitor);
                if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor)
                    output = new HtmlParser().ParseDocument(output).Body!.TextContent;
                Assert.Equal(string.Concat(expected.Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    public static string[] Labels(int depth, string mode, bool root)
    {
        if (!root && mode is "none" or "start") return Enumerable.Repeat("", 6).ToArray();
        if (depth > 0) return ["", "BEGIN7END", "BEGIN8END", "", "BEGIN7END", "BEGIN8END"];
        return Enumerable.Range(mode is "start" or "both" ? 7 : root ? 2 : 0, 6).Select(n => "BEGIN" + n + "END").ToArray();
    }

    public static byte[] CreateSource(int depth, string mode, bool root)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            if (!root) document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<AbstractNum>().Single().RemoveAllChildren<Level>();
            if (depth > 0)
            {
                var levels = new[] { 0, depth, depth, 0, depth, depth };
                foreach (var (paragraph, i) in document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Select((p, i) => (p, i)))
                    paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val = levels[i];
            }
        }
        return stream.ToArray();
    }
}
