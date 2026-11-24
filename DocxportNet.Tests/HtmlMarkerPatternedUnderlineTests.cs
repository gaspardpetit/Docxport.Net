using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class HtmlMarkerPatternedUnderlineTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlMarkerPatternedUnderline");
    }

    public static IEnumerable<object?[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return [row.GetProperty("case").GetString()!,
                row.GetProperty("expectedStrikeLines").GetInt32(), row.GetProperty("expectedUnderlineStyle").GetString()!,
                row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(label => label.GetString()!).ToArray()];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void PreservesPatternedUnderlineWithIndependentStrike(string name, int expectedStrikeLines, string expectedUnderlineStyle, string[] labels)
    {
        var path = Path.Combine(FixtureRoot(), "source", name + ".docx");
        using var document = WordprocessingDocument.Open(path, false);
        using (var validation = WordprocessingDocument.Open(path, false, new OpenSettings {
            MarkupCompatibilityProcessSettings = new(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013) }))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(validation).ToArray();
            Assert.Empty(errors);
        }
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        foreach (var rich in new[] { true, false })
        {
            var html = DxpExport.ExportToString(document, new DxpHtmlVisitor(rich
                ? DxpHtmlVisitorConfig.CreateRichConfig() : DxpHtmlVisitorConfig.CreatePlainConfig()));
            var markers = new HtmlParser().ParseDocument(html).QuerySelectorAll(".dxp-marker");
            Assert.Equal(labels, markers.Select(marker => marker.TextContent.Trim()));
            foreach (var marker in markers)
            {
                var css = marker.GetAttribute("style") ?? "";
                var doubles = marker.QuerySelectorAll(".dxp-marker-double-strike");
                Assert.Contains("text-decoration:underline", css);
                Assert.Contains("text-decoration-style:" + expectedUnderlineStyle, css);
                Assert.DoesNotContain("line-through", css);
                var singles = marker.QuerySelectorAll(".dxp-marker-single-strike");
                if (expectedStrikeLines == 1)
                    Assert.Equal("text-decoration:line-through;text-decoration-style:solid", Assert.Single(singles).GetAttribute("style"));
                else Assert.Empty(singles);
                if (expectedStrikeLines == 2)
                    Assert.Equal("text-decoration:line-through;text-decoration-style:double",
                        Assert.Single(doubles).GetAttribute("style"));
                else Assert.Empty(doubles);
            }
        }
        Assert.Equal(before, Snapshot());
    }
}
