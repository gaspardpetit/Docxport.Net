using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Html;

namespace DocxportNet.Tests;

public sealed class MarkerNumericParagraphRtlSizeTests
{
    public static IEnumerable<object[]> Cases()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        if (root == null) throw new DirectoryNotFoundException("Repository fixtures not found.");
        var folder = Path.Combine(root.FullName, "fixtures/docx/regressions/list-markers/NumericRtlMarkerSize/source");
        foreach (var path in Directory.GetFiles(folder, "*.docx").OrderBy(path => path))
        {
            var name = Path.GetFileNameWithoutExtension(path);
            using var document = WordprocessingDocument.Open(path, false);
            var main = document.MainDocumentPart!;
            var reset = new[] { "direct-off", "style-reset", "paragraph-reset", "defaults-reset" }.Any(name.StartsWith);
            yield return new object[] { name, main.StyleDefinitionsPart!.Styles!.OuterXml, main.Document!.OuterXml,
                main.NumberingDefinitionsPart!.Numbering!.OuterXml, reset ? 24 : 30, name.EndsWith("literal") ? "12" : "1" };
        }
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void MatchesWordNumericMarkerSize(string name, string stylesXml, string documentXml, string numberingXml, int expectedSize, string label)
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Document(documentXml);
        main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(stylesXml);
        main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(numberingXml);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        _ = main.Document.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        var before = new[] { main.Document!.OuterXml, main.StyleDefinitionsPart!.Styles!.OuterXml, main.NumberingDefinitionsPart!.Numbering!.OuterXml };
        var html = DxpExport.ExportToString(document, new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()));
        var markers = new HtmlParser().ParseDocument(html).QuerySelectorAll(".dxp-marker");
        Assert.True(markers.Length == 3, name);
        Assert.Equal(label == "12" ? new[] { "12", "12", "12" } : new[] { "1", "2", "3" }, markers.Select(marker => marker.TextContent.Trim()));
        foreach (var marker in markers)
        {
            if (expectedSize == 30) Assert.Contains("font-size:15pt", marker.GetAttribute("style"));
            else Assert.DoesNotContain("font-size:15pt", marker.OuterHtml);
            if (expectedSize == 24) Assert.DoesNotContain("font-size:15pt", marker.InnerHtml);
        }
        var plain = DxpExport.ExportToString(document, new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()));
        Assert.All(new HtmlParser().ParseDocument(plain).QuerySelectorAll(".dxp-marker"),
            marker => Assert.DoesNotContain("font-size:15pt", marker.OuterHtml));
        Assert.Equal(before, new[] { main.Document!.OuterXml, main.StyleDefinitionsPart!.Styles!.OuterXml, main.NumberingDefinitionsPart!.Numbering!.OuterXml });
    }
}
