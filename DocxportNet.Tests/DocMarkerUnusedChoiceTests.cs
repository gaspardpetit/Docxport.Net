using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocMarkerUnusedChoiceTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var scope in new[] { "style", "properties", "numbering", "id", "default" })
                foreach (var branch in new[] { "choice", "fallback", "direct" })
                    yield return new object[] { depth, scope, branch };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void RejectedStyleNumberingDoesNotBlockDocExport(int depth, string scope, string branch)
    {
        using var input = new MemoryStream(CreateSource(depth, scope, branch));
        using var source = WordprocessingDocument.Open(input, false);
        var main = source.MainDocumentPart!;
        _ = main.Document.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        var before = new[] { main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml, main.NumberingDefinitionsPart.Numbering.OuterXml };
        var binary = DxpExport.ExportToBytes(source, new DxpDocVisitor(), options: null);
        Assert.Equal(before, new[] { main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml, main.NumberingDefinitionsPart.Numbering.OuterXml });
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(projected));
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
                Assert.Equal("[7]Item0[8]Item1[9]Item2", Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    [Fact]
    public void SelectedUnsupportedNumberingStillFails()
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource(0, "style", "choice"));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
            document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<AbstractNum>().Single(x => x.AbstractNumberId!.Value == 1)
                .GetFirstChild<Level>()!.NumberingFormat!.Val = NumberFormatValues.Hex;
        Assert.Throws<NotSupportedException>(() => DocxportNet.DxpDocExport.Export(stream.ToArray()));
    }

    [Fact]
    public void UnusedAuthoredStyleStillRequiresItsListDefinition()
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource(0, "style", "direct"));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
            document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new Style(
                new StyleName { Val = "Unused" }, new StyleParagraphProperties(new NumberingProperties(
                    new NumberingId { Val = 2 }))) { StyleId = "Unused", Type = StyleValues.Paragraph });
        Assert.Throws<NotSupportedException>(() => DocxportNet.DxpDocExport.Export(stream.ToArray()));
    }

    public static byte[] CreateSource(int depth, string scope, string branch)
    {
        var bytes = DocMarkerStyleChoiceTests.CreateSource(depth, scope, branch);
        using var edited = new MemoryStream();
        edited.Write(bytes);
        edited.Position = 0;
        using (var document = WordprocessingDocument.Open(edited, true))
        {
            var numbering = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            numbering.Elements<AbstractNum>().Single(x => x.AbstractNumberId!.Value == 2)
                .GetFirstChild<Level>()!.NumberingFormat!.Val = NumberFormatValues.Hex;
        }
        return edited.ToArray();
    }
}
