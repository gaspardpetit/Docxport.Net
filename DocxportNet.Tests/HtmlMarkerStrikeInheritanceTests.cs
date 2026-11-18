using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using System.Globalization;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class HtmlMarkerStrikeInheritanceTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlMarkerStrikeInheritance");
    }

    public static IEnumerable<object?[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return [row.GetProperty("case").GetString()!,
                row.GetProperty("expectedStrike").GetBoolean(),
                row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(label => label.GetString()!).ToArray()];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void InheritsSingleStrikeWithoutChangingBodyFormatting(string name, bool expectedStrike, string[] labels)
    {
        var path = Path.Combine(FixtureRoot(), "source", name + ".docx");
        using var document = WordprocessingDocument.Open(path, false);
        using (var validation = WordprocessingDocument.Open(path, false, new OpenSettings {
            MarkupCompatibilityProcessSettings = new(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013) }))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(validation).ToArray();
            if (name.Contains("-stylemark-", StringComparison.Ordinal))
            {
                Assert.NotEmpty(errors);
                foreach (var error in errors)
                {
                    Assert.Equal("Sch_InvalidElementContentExpectingComplex", error.Id);
                    Assert.Equal("/word/styles.xml", error.Part!.Uri.ToString());
                    Assert.Equal("pPr", error.Node!.LocalName);
                    Assert.Equal("style", error.Node.Parent!.LocalName);
                    Assert.Contains(":rPr", error.Description);
                }
            }
            else Assert.Empty(errors);
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
                if (expectedStrike)
                    Assert.Contains("line-through", marker.GetAttribute("style") ?? "");
                else Assert.DoesNotContain("line-through", marker.GetAttribute("style") ?? "");
        }
        Assert.Equal(before, Snapshot());
    }
}
