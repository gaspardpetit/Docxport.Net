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

public sealed class DocMarkerStoryChoiceDiscoveryTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var story in new[] { "body", "header", "footer" })
                foreach (var branch in new[] { "direct", "fallback", "nested" })
                    yield return new object[] { depth, story, branch };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void RejectedStoryParagraphsDoNotBlockDocExport(int depth, string story, string branch)
    {
        using var input = new MemoryStream(CreateSource(depth, story, branch));
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
                Assert.True(expected == Regex.Replace(output, @"\s+", ""), $"{(ReferenceEquals(document, source) ? "source" : "projected")} {visitor.GetType().Name}: {output}");
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
        stream.Write(CreateSource(0, story, "fallback"));
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
            document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<AbstractNum>().Single(x => x.AbstractNumberId!.Value == 1)
                .GetFirstChild<Level>()!.NumberingFormat!.Val = NumberFormatValues.Hex;
        Assert.Throws<NotSupportedException>(() => DocxportNet.DxpDocExport.Export(stream.ToArray()));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void DiscoveryUsesTheWritersChoicePolicy(bool acceptChoice, bool swapBranches)
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource(0, "body", "fallback"));
        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, true);
        var alternate = document.MainDocumentPart!.Document.Descendants<AlternateContent>().Single();
        if (swapBranches)
        {
            var choice = alternate.GetFirstChild<AlternateContentChoice>()!;
            var fallback = alternate.GetFirstChild<AlternateContentFallback>()!;
            var selected = fallback.ChildElements.Select(child => child.CloneNode(true)).ToArray();
            var rejected = choice.ChildElements.Select(child => child.CloneNode(true)).ToArray();
            choice.RemoveAllChildren(); fallback.RemoveAllChildren();
            choice.Append(selected); fallback.Append(rejected);
        }
        var before = document.MainDocumentPart.Document.OuterXml;
        (DocxportNet.Doc.DocListIndex Lists, IReadOnlyDictionary<int, short> NumberIds) Read() =>
            DocxportNet.Doc.DocOpenXmlListReader.Read(document.MainDocumentPart,
                _ => DocxportNet.Doc.DocCharacterFormatting.Empty,
                _ => DocxportNet.Doc.DocParagraphFormatting.Empty, _ => acceptChoice);
        if (acceptChoice != swapBranches) Assert.Throws<NotSupportedException>(() => Read());
        else Assert.Equal(1, Assert.Single(Read().NumberIds).Key);
        Assert.Equal(before, document.MainDocumentPart.Document.OuterXml);
    }

    public static byte[] CreateSource(int depth, string story, string branch)
    {
        var bytes = DocMarkerUnusedParagraphChoiceTests.CreateSource(depth, "properties", "direct", story);
        if (branch == "direct") return bytes;
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            OpenXmlCompositeElement root = story == "header" ? main.HeaderParts.Single().Header! :
                story == "footer" ? main.FooterParts.Single().Footer! : main.Document.Body!;
            if (!root.NamespaceDeclarations.Any(pair => pair.Key == "mc")) root.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            if (!root.NamespaceDeclarations.Any(pair => pair.Key == "x")) root.AddNamespaceDeclaration("x", "urn:marker-unsupported");
            root.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "x" };
            var selected = root.Elements<Paragraph>().ToArray();
            var rejected = selected.Select(paragraph =>
            {
                var clone = (Paragraph)paragraph.CloneNode(true);
                clone.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 2;
                return clone;
            }).ToArray();
            foreach (var paragraph in selected) paragraph.Remove();
            AlternateContent Wrap(IEnumerable<OpenXmlElement> children) => new(
                new AlternateContentChoice(rejected.Select(p => p.CloneNode(true))) { Requires = "x" },
                new AlternateContentFallback(children));
            var alternate = Wrap(selected);
            if (branch == "nested") alternate = Wrap(new[] { alternate });
            root.PrependChild(alternate);
        }
        return stream.ToArray();
    }
}
