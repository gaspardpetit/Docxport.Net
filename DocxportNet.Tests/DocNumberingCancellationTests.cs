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

public sealed class DocNumberingCancellationTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 1, 8 })
            foreach (var mode in new[] { "direct", "direct-level", "style", "control", "reenabled", "resume" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void CancellationSurvivesBinaryExportAndProjection(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Equal(9, Assert.Single(index.Lists.Definitions).Levels.Count);
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
        mode == "resume" ? ["BEGIN2END", "BEGIN5END", "", "BEGIN3END", "", "BEGIN5END"] :
        mode is "control" or "reenabled"
            ? ["BEGIN2END", "BEGIN5END", "BEGIN6END", "BEGIN3END", "BEGIN5END", "BEGIN6END"]
            : ["BEGIN2END", "", "", "BEGIN3END", "", ""];

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocStyleLevelContextTests.CreateSource(depth, "full-style"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            if (mode is "style" or "reenabled")
                main.StyleDefinitionsPart!.Styles!.AppendChild(new Style(
                    new StyleName { Val = "Disabled" }, new BasedOn { Val = "Nested" },
                    new StyleParagraphProperties(new NumberingProperties(new NumberingId { Val = 0 })))
                    { Type = StyleValues.Paragraph, StyleId = "Disabled" });
            foreach (var (paragraph, i) in main.Document.Body!.Elements<Paragraph>().Select((p, i) => (p, i)))
            {
                if (i is 0 or 3) continue;
                var properties = paragraph.ParagraphProperties!;
                if (mode is "style" or "reenabled")
                    properties.ParagraphStyleId = new ParagraphStyleId { Val = "Disabled" };
                properties.NumberingProperties = mode switch
                {
                    "resume" when i is 2 or 4 => new NumberingProperties(new NumberingId { Val = 0 }),
                    "direct" => new NumberingProperties(new NumberingId { Val = 0 }),
                    "direct-level" => new NumberingProperties(new NumberingLevelReference { Val = depth }, new NumberingId { Val = 0 }),
                    "reenabled" => new NumberingProperties(new NumberingLevelReference { Val = depth }, new NumberingId { Val = 1 }),
                    _ => null
                };
            }
        }
        return stream.ToArray();
    }
}
