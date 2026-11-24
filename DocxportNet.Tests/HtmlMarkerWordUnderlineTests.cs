using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class HtmlMarkerWordUnderlineTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlMarkerWordUnderline");
    }

    public static IEnumerable<object[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return [row.GetProperty("case").GetString()!, row.GetProperty("expectedStrikeLines").GetInt32(),
                row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(v => v.GetString()!).ToArray(),
                row.GetProperty("expectedUnderlinedWords").EnumerateArray().Select(words =>
                    words.EnumerateArray().Select(v => v.GetString()!).ToArray()).ToArray()];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void UnderlinesWordSegmentsWithoutChangingLiteralText(string name, int strikeLines, string[] labels, string[][] words)
    {
        var path = Path.Combine(FixtureRoot(), "source", name + ".docx");
        using var document = WordprocessingDocument.Open(path, false);
        using (var validation = WordprocessingDocument.Open(path, false, new OpenSettings {
            MarkupCompatibilityProcessSettings = new(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013) }))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(validation));
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml, main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        foreach (var rich in new[] { true, false })
        {
            var html = DxpExport.ExportToString(document, new DxpHtmlVisitor(rich
                ? DxpHtmlVisitorConfig.CreateRichConfig() : DxpHtmlVisitorConfig.CreatePlainConfig()));
            var markers = new HtmlParser().ParseDocument(html).QuerySelectorAll(".dxp-marker");
            Assert.Equal(labels, markers.Select(m => m.TextContent.Trim()));
            for (var i = 0; i < markers.Length; i++)
            {
                var marker = markers[i];
                var css = marker.GetAttribute("style") ?? "";
                Assert.DoesNotContain("underline", css);
                Assert.Equal(strikeLines == 1, css.Contains("line-through", StringComparison.Ordinal));
                var segments = marker.QuerySelectorAll(".dxp-marker-word-underline");
                Assert.Equal(words[i], segments.Select(s => s.TextContent));
                Assert.All(segments, s => Assert.Equal("text-decoration:underline", s.GetAttribute("style")));
                Assert.Empty(marker.QuerySelectorAll("script,b"));
                Assert.Empty(marker.QuerySelectorAll(".dxp-marker-single-strike"));
                var doubles = marker.QuerySelectorAll(".dxp-marker-double-strike");
                if (strikeLines == 2)
                    Assert.Equal("text-decoration:line-through;text-decoration-style:double", Assert.Single(doubles).GetAttribute("style"));
                else Assert.Empty(doubles);
            }
        }
        Assert.Equal(before, Snapshot());
    }
}
