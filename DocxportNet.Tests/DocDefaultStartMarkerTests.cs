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

public sealed class DocDefaultStartMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 1, 8 })
            foreach (int? start in new int?[] { null, 0, 1, 7 })
                yield return new object[] { depth, start!, false };
        yield return new object[] { 0, 1, true };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void OmittedStartMatchesZeroAcrossEditableDocRoundTrip(int depth, int? start, bool replacement)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var levels = Enumerable.Range(0, depth + 1).Select(index =>
            {
                var level = new Level { LevelIndex = index };
                if (index != depth || start != null)
                    level.AppendChild(new StartNumberingValue { Val = index == depth ? start!.Value : 3 });
                level.AppendChild(new NumberingFormat { Val = NumberFormatValues.Decimal });
                level.AppendChild(new LevelText { Val = (index == depth ? "BEGIN" : "ROOT") + "%" + (index + 1) + "END" });
                return level;
            });
            var number = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 };
            if (replacement)
                number.AppendChild(new LevelOverride(new Level(new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "BEGIN%1END" }) { LevelIndex = 0 }) { LevelIndex = 0 });
            var numbering = new Numbering(new AbstractNum(levels) { AbstractNumberId = 1 }, number);
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = numbering;
            var usedLevels = depth == 0 ? new[] { 0, 0, 0, 0, 0, 0 } : new[] { 0, depth, depth, 0, depth, depth };
            main.Document = new Document(new Body(usedLevels.Select((level, i) =>
                new Paragraph(new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = level }, new NumberingId { Val = 1 })),
                    new Run(new Text("Item" + i))))));
            var original = numbering.OuterXml;
            _ = DocOpenXmlListReader.Read(main, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
            Assert.Equal(original, numbering.OuterXml);
        }
        var first = start ?? 0;
        var expected = depth == 0 ? Enumerable.Range(first, 6).Select(n => "BEGIN" + n + "END").ToArray() :
            new[] { "ROOT3END", "BEGIN" + first + "END", "BEGIN" + (first + 1) + "END",
                "ROOT4END", "BEGIN" + first + "END", "BEGIN" + (first + 1) + "END" };
        var binary = DxpDocExport.Export(stream.ToArray());
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
            Assert.Equal(first, index.Lists.Definitions.Single().Levels.Single(level => level.Index == depth).StartAt);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        using var sourceStream = new MemoryStream(stream.ToArray());
        using var source = WordprocessingDocument.Open(sourceStream, false);
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
                Assert.Equal(expected, Regex.Matches(output, "(?:BEGIN|ROOT).*?END").Select(m => m.Value));
                Assert.Equal(6, Regex.Matches(output, @"Item\d+").Count);
            }
        }
    }
}
