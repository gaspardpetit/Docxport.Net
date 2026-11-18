using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class HtmlLiteralMarkerMarkupTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlLiteralMarkerMarkup");
    }

    public static IEnumerable<object[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return [row.GetProperty("case").GetString()!,
                row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(label => label.GetString()!).ToArray(),
                row.GetProperty("generatedFontWrapper").GetBoolean()];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void PreservesLiteralLabelsAndGeneratedFontWrappers(string name, string[] labels, bool fontWrapper)
    {
        using var document = WordprocessingDocument.Open(Path.Combine(FixtureRoot(), "source", name + ".docx"), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        foreach (var rich in new[] { true, false })
        {
            var output = DxpExport.ExportToString(document, new DxpHtmlVisitor(rich
                ? DxpHtmlVisitorConfig.CreateRichConfig() : DxpHtmlVisitorConfig.CreatePlainConfig()));
            var markers = new HtmlParser().ParseDocument(output).QuerySelectorAll(".dxp-marker");
            Assert.Equal(labels, markers.Select(marker => marker.TextContent.Trim()));
            foreach (var marker in markers)
            {
                var elements = marker.QuerySelectorAll("*");
                if (rich && fontWrapper)
                {
                    var wrapper = Assert.Single(elements);
                    Assert.Equal("SPAN", wrapper.TagName);
                    Assert.Equal("font-family:Arial", wrapper.GetAttribute("style"));
                }
                else Assert.Empty(elements);
            }
        }
        Assert.Equal(before, Snapshot());
    }
}
