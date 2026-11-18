using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using System.Globalization;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class HtmlNormalMarkerSizeInheritanceTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlNormalMarkerSizeInheritance");
    }

    public static IEnumerable<object?[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return [row.GetProperty("case").GetString()!,
                row.GetProperty("expectedHalfPoints").ValueKind == JsonValueKind.Null ? null : row.GetProperty("expectedHalfPoints").GetInt32(),
                row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(label => label.GetString()!).ToArray()];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void InheritsNormalSizeWithoutChangingScriptContexts(string name, int? expectedHalfPoints, string[] labels)
    {
        var path = Path.Combine(FixtureRoot(), "source", name + ".docx");
        using var document = WordprocessingDocument.Open(path, false);
        using (var validation = WordprocessingDocument.Open(path, false, new OpenSettings {
            MarkupCompatibilityProcessSettings = new(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013) }))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2013).Validate(validation).ToArray();
            if (name is "guard-style-cs" or "guard-defaults-cs")
            {
                // Word accepts cs in these locations; the SDK schema does not.
                Assert.NotEmpty(errors);
                foreach (var error in errors)
                {
                    Assert.Equal("Sch_InvalidElementContentExpectingComplex", error.Id);
                    Assert.Equal("/word/styles.xml", error.Part!.Uri.ToString());
                    Assert.Equal("rPr", error.Node!.LocalName);
                    Assert.Equal(name == "guard-style-cs" ? "style" : "rPrDefault", error.Node.Parent!.LocalName);
                    Assert.Contains(":cs", error.Description);
                }
            }
            else if (name.StartsWith("style-mark-", StringComparison.Ordinal))
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
                if (rich && expectedHalfPoints is int halfPoints)
                    Assert.Contains("font-size:" + (halfPoints / 2.0).ToString("0.###", CultureInfo.InvariantCulture) + "pt", marker.GetAttribute("style"));
                else Assert.DoesNotContain("font-size:", marker.GetAttribute("style") ?? "");
        }
        Assert.Equal(before, Snapshot());
    }
}
