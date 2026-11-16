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

public sealed class DocMarkerUnusedParagraphChoiceTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var scope in new[] { "properties", "numbering", "id" })
                foreach (var story in new[] { "body", "header", "footer" })
                    foreach (var branch in new[] { "choice", "fallback", "direct" })
                        yield return new object[] { depth, scope, branch, story };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void RejectedParagraphNumberingDoesNotBlockDocExport(int depth, string scope, string branch, string story)
    {
        using var input = new MemoryStream(CreateSource(depth, scope, branch, story));
        using var source = WordprocessingDocument.Open(input, false);
        var main = source.MainDocumentPart!;
        _ = main.Document.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => new[] { main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml }.Concat(main.HeaderParts.Select(p => p.Header!.OuterXml))
            .Concat(main.FooterParts.Select(p => p.Footer!.OuterXml)).ToArray();
        var before = Snapshot();
        var binary = DxpExport.ExportToBytes(source, new DxpDocVisitor(), options: null);
        Assert.Equal(before, Snapshot());
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
                // PlainText intentionally suppresses header/footer stories.
                var expected = story != "body" && visitor is DxpPlainTextVisitor
                    ? "" : "[7]Item0[8]Item1[9]Item2";
                Assert.Equal(expected, Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void SelectedUnsupportedNumberingStillFails(string story)
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource(0, "properties", "choice", story));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
            document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<AbstractNum>().Single(x => x.AbstractNumberId!.Value == 1)
                .GetFirstChild<Level>()!.NumberingFormat!.Val = NumberFormatValues.Hex;
        Assert.Throws<NotSupportedException>(() => DocxportNet.DxpDocExport.Export(stream.ToArray()));
    }

    public static byte[] CreateSource(int depth, string scope, string branch, string story)
    {
        var bytes = DocMarkerPropertyChoiceTests.CreateSource(depth, scope, branch);
        using var edited = new MemoryStream();
        edited.Write(bytes);
        edited.Position = 0;
        using (var document = WordprocessingDocument.Open(edited, true))
        {
            var main = document.MainDocumentPart!;
            // The rejected list must not also be referenced by an ordinary style.
            main.StyleDefinitionsPart!.Styles!.Elements<Style>().Single(x => x.StyleId == "Wrong").Remove();
            if (story != "body")
            {
                var paragraphs = main.Document.Body!.Elements<Paragraph>().ToArray();
                foreach (var paragraph in paragraphs) paragraph.Remove();
                main.Document.Body.Append(new Paragraph());
                var section = new SectionProperties();
                if (story == "header")
                {
                    var part = main.AddNewPart<HeaderPart>();
                    part.Header = new Header(paragraphs);
                    part.Header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                    part.Header.AddNamespaceDeclaration("x", "urn:marker-unsupported");
                    part.Header.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x" };
                    section.Append(new HeaderReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default });
                }
                else
                {
                    var part = main.AddNewPart<FooterPart>();
                    part.Footer = new Footer(paragraphs);
                    part.Footer.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                    part.Footer.AddNamespaceDeclaration("x", "urn:marker-unsupported");
                    part.Footer.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x" };
                    section.Append(new FooterReference { Id = main.GetIdOfPart(part), Type = HeaderFooterValues.Default });
                }
                main.Document.Body.Append(section);
            }
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            numbering.Elements<AbstractNum>().Single(x => x.AbstractNumberId!.Value == 2)
                .GetFirstChild<Level>()!.NumberingFormat!.Val = NumberFormatValues.Hex;
        }
        return edited.ToArray();
    }
}
