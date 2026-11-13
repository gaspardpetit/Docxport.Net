using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocSparseAbstractLevelTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 1, 2, 8 })
        {
            var body = new[] { 0, depth, depth, 0, depth, depth };
            var child = new[] { "BEGIN3END", "BEGIN4END" };
            yield return new object[] { "dense-" + depth, Enumerable.Range(0, depth + 1).ToArray(), body, false,
                new[] { "ROOT3END", child[0], child[1], "ROOT4END", child[0], child[1] } };
            yield return new object[] { "missing-root-" + depth, Enumerable.Range(1, depth).ToArray(), body, false,
                new[] { "", child[0], child[1], "", child[0], child[1] } };
            if (depth == 1) continue;
            yield return new object[] { "only-level-" + depth, new[] { depth }, body, false,
                new[] { "", child[0], child[1], "", child[0], child[1] } };
            yield return new object[] { "gap-" + depth, new[] { 0, depth }, new[] { 0, 1, depth, 0, 1, depth }, false,
                new[] { "ROOT3END", "", child[0], "ROOT4END", "", child[0] } };
            yield return new object[] { "ancestor-gap-" + depth, new[] { 0, depth }, body, true,
                new[] { "ROOT3END", "BEGIN3/0/3END", "BEGIN3/0/4END", "ROOT4END", "BEGIN4/0/3END", "BEGIN4/0/4END" } };
        }
        yield return new object[] { "missing-tail-2", new[] { 0, 1 }, new[] { 0, 2, 2, 0, 2, 2 }, false,
            new[] { "ROOT3END", "", "", "ROOT4END", "", "" } };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void MissingAbstractSlotsStayEmptyAndStartAtZero(string name, int[] defined, int[] body,
        bool ancestor, string[] expected)
    {
        var bytes = CreateSource(defined, body, ancestor);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var levels = Assert.Single(index.Lists.Definitions).Levels;
            Assert.Equal(Enumerable.Range(0, 9), levels.Select(x => x.Index));
            foreach (var level in levels.Where(x => !defined.Contains(x.Index)))
            {
                Assert.Equal(0, level.StartAt);
                Assert.Equal("", level.NumberText);
                Assert.Equal(0, level.NumberFormat);
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
                Assert.Equal(string.Concat(expected.Select((marker, i) => marker + "Item" + i)),
                    Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    public static byte[] CreateSource(int[] defined, int[] body, bool ancestor)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var depth = defined.Max();
            var numbering = new Numbering(new AbstractNum(defined.Select(i => new Level(
                new StartNumberingValue { Val = 3 }, new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = i == 0 ? "ROOT%1END" : ancestor && i == depth ? "BEGIN%1/%2/%" + (i + 1) + "END" : "BEGIN%" + (i + 1) + "END" })
                { LevelIndex = i })) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
            main.Document = new Document(new Body(body.Select((level, i) =>
                new Paragraph(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = level }, new NumberingId { Val = 1 })),
                    new Run(new Text("Item" + i))))));
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(9)]
    public void InvalidIndexesRemainRejected(int index) =>
        Assert.Throws<InvalidDataException>(() => DxpDocExport.Export(CreateSource([index], [0], false)));

    [Fact]
    public void DuplicateIndexesRemainRejected() =>
        Assert.Throws<NotSupportedException>(() => DxpDocExport.Export(CreateSource([0, 0], [0], false)));
}
