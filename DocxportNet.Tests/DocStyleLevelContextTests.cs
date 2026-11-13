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

public sealed class DocStyleLevelContextTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 1, 8 })
            foreach (var mode in new[] { "inherited-id", "inherited-level", "full-style" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void StyleMembershipPreservesInheritedLevels(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var levels = Assert.Single(index.Lists.Definitions).Levels;
            Assert.Equal(9, levels.Count);
            foreach (var level in levels)
            {
                Assert.Equal(level.Index == 0 ? 2 : level.Index == depth && mode == "full-style" ? 5 : 0, level.StartAt);
                Assert.Equal(level.Index == 0 ? "BEGIN%1END" : level.Index == depth && mode == "full-style" ? "BEGIN%" + (depth + 1) + "END" : "", level.NumberText);
            }
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        var styles = source.MainDocumentPart.StyleDefinitionsPart!.Styles!;
        var originalStyles = styles.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        Assert.Equal(originalStyles, styles.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
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

    public static string[] Labels(int depth, string mode) =>
        mode == "full-style"
            ? ["BEGIN2END", "BEGIN5END", "BEGIN6END", "BEGIN3END", "BEGIN5END", "BEGIN6END"]
            : ["BEGIN2END", "", "", "BEGIN3END", "", ""];

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocRootHigherLevelTests.CreateSource(depth, "none", false));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            if (mode == "full-style")
                main.NumberingDefinitionsPart!.Numbering!.Elements<AbstractNum>().Single().AppendChild(
                    new Level(new StartNumberingValue { Val = 5 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = "BEGIN%" + (depth + 1) + "END" }) { LevelIndex = depth });
            var styleNumbering = new NumberingProperties();
            if (mode != "inherited-level") styleNumbering.NumberingId = new NumberingId { Val = 1 };
            if (mode != "inherited-id") styleNumbering.NumberingLevelReference = new NumberingLevelReference { Val = depth };
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new Style(
                new StyleName { Val = "Nested" }, new StyleParagraphProperties(styleNumbering))
                { Type = StyleValues.Paragraph, StyleId = "Nested" });
            foreach (var (paragraph, i) in main.Document.Body!.Elements<Paragraph>().Select((p, i) => (p, i)))
            {
                if (i is 0 or 3) continue;
                var properties = paragraph.ParagraphProperties!;
                properties.ParagraphStyleId = new ParagraphStyleId { Val = "Nested" };
                properties.NumberingProperties = mode switch
                {
                    "inherited-id" => new NumberingProperties(new NumberingLevelReference { Val = depth }),
                    "inherited-level" => new NumberingProperties(new NumberingId { Val = 1 }),
                    _ => null
                };
            }
        }
        return stream.ToArray();
    }
}
