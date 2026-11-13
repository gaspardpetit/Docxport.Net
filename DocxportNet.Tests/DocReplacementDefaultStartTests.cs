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

public sealed class DocReplacementDefaultStartTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (var start in new[] { "missing", "zero", "seven" })
                foreach (var context in depth == 0 ? new[] { "repeat" } : new[] { "parent", "repeat" })
                    yield return new object[] { depth, start + "-" + context };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void OmittedReplacementStartPreservesWordSequence(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Single(index.Lists.Definitions);
            var instance = Assert.Single(index.Lists.Instances);
            Assert.Empty(instance.StartOverrides);
            Assert.Equal(mode.StartsWith("seven") ? 7 : mode.StartsWith("zero") || depth > 0 ? 0 : 1,
                instance.FormattingOverrides[depth].StartAt);
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
        var projectedNumbering = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        Assert.Equal(1, projectedNumbering.Elements<AbstractNum>().Count());
        Assert.Single(projectedNumbering.Elements<NumberingInstance>());
        Assert.Equal(1,
            projectedNumbering.Descendants<LevelOverride>().Count(x => x.Level != null));
        var expected = Labels(depth, mode);
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

    public static string[] Labels(int depth, string mode)
    {
        var start = mode.StartsWith("seven") ? 7 : 0;
        var values = mode.EndsWith("parent")
            ? new[] { 3, start, start + 1, 4, start, start + 1 }
            : Enumerable.Range(3, 6).ToArray();
        return values.Select(value => "BEGIN" + value + "END").ToArray();
    }

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            foreach (var level in numbering.Descendants<Level>()) level.StartNumberingValue!.Val = 3;
            var replacement = new Level(new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "BEGIN%" + (depth + 1) + "END" }) { LevelIndex = depth };
            if (!mode.StartsWith("missing"))
                replacement.PrependChild(new StartNumberingValue { Val = mode.StartsWith("seven") ? 7 : 0 });
            numbering.Elements<NumberingInstance>().Single().AppendChild(
                new LevelOverride(replacement) { LevelIndex = depth });
            if (mode.EndsWith("parent"))
                foreach (var (paragraph, i) in main.Document.Body!.Elements<Paragraph>().Select((p, i) => (p, i)))
                    if (i is 0 or 3) paragraph.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val = depth - 1;
        }
        return stream.ToArray();
    }
}
