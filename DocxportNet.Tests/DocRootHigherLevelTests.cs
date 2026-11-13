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

public sealed class DocRootHigherLevelTests
{
    [Fact]
    public void UnusedHigherOverrideDoesNotExpandSingleRootDefinition()
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource(0, "none", true));
        using var document = WordprocessingDocument.Open(stream, true);
        document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.AppendChild(
            new NumberingInstance(new AbstractNumId { Val = 1 },
                new LevelOverride(new StartOverrideNumberingValue { Val = 7 }) { LevelIndex = 8 })
                { NumberID = 2 });
        var result = DocOpenXmlListReader.Read(document.MainDocumentPart,
            _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Single(Assert.Single(result.Lists.Definitions).Levels);
        Assert.Single(result.Lists.Instances);
    }

    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 1, 8 })
            foreach (var mode in new[] { "none", "start", "format", "both" })
                yield return new object[] { depth, mode, false };
        foreach (var mode in new[] { "none", "start", "format", "both" })
            yield return new object[] { 0, mode, true };
        foreach (var depth in new[] { 1, 8 })
            foreach (var mode in new[] { "start", "format", "both" })
                yield return new object[] { depth, mode, true };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void RootDefinitionsPreserveHigherSlotsAndReplacementActivation(int depth, string mode, bool root)
    {
        var bytes = CreateSource(depth, mode, root);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var levels = Assert.Single(index.Lists.Definitions).Levels;
            Assert.Equal(depth == 0 ? 1 : 9, levels.Count);
            foreach (var level in levels)
            {
                Assert.Equal(level.Index == 0 ? 2 : 0, level.StartAt);
                Assert.Equal(level.Index == 0 ? "BEGIN%1END" : "", level.NumberText);
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
        if (depth > 0 && !root) return mode is "format" or "both"
            ? ["BEGIN2END", "BEGIN7END", "BEGIN8END", "BEGIN3END", "BEGIN7END", "BEGIN8END"]
            : ["BEGIN2END", "", "", "BEGIN3END", "", ""];
        return Enumerable.Range(depth == 0 && mode is "start" or "both" ? 7 : 2, 6).Select(n => "BEGIN" + n + "END").ToArray();
    }

    public static byte[] CreateSource(int depth, string mode, bool root)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            foreach (var level in document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<AbstractNum>().Single().Elements<Level>().Where(l => l.LevelIndex!.Value > 0).ToArray()) level.Remove();
            if (root)
                foreach (var paragraph in document.MainDocumentPart!.Document.Body!.Elements<Paragraph>())
                    paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val = 0;
            if (depth > 0 && !root)
            {
                var levels = new[] { 0, depth, depth, 0, depth, depth };
                foreach (var (paragraph, i) in document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().Select((p, i) => (p, i)))
                    paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val = levels[i];
            }
        }
        return stream.ToArray();
    }
}
