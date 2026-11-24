using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class HtmlMarkerUnderlineThemeBaseTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlMarkerUnderlineThemeBase");
    }

    public static IEnumerable<object?[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return [row.GetProperty("case").GetString()!, row.GetProperty("underlineStyle").GetString()!,
                row.GetProperty("expectedCssColor").GetString(), row.GetProperty("expectedStrikeLines").GetInt32(),
                row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(v => v.GetString()!).ToArray()];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ResolvesBaseThemeWithoutApplyingCachedRgb(string name, string style, string? expectedColor, int strikeLines, string[] labels)
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
        _ = main.ThemePart?.Theme?.Descendants().ToArray();
        _ = main.DocumentSettingsPart?.Settings?.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml, main.ThemePart?.Theme?.OuterXml ?? "", main.DocumentSettingsPart?.Settings?.OuterXml ?? ""];
        var before = Snapshot();
        var configs = new[] {
            (DxpHtmlVisitorConfig.CreateRichConfig(), true), (DxpHtmlVisitorConfig.CreatePlainConfig(), false),
            (DxpHtmlVisitorConfig.CreateRichConfig() with { EmitRunColor = false }, false),
            (DxpHtmlVisitorConfig.CreatePlainConfig() with { EmitRunColor = true }, true) };
        foreach (var (config, emitColor) in configs)
        {
            var html = DxpExport.ExportToString(document, new DxpHtmlVisitor(config));
            var markers = new HtmlParser().ParseDocument(html).QuerySelectorAll(".dxp-marker");
            Assert.Equal(labels, markers.Select(m => m.TextContent.Trim()));
            var color = emitColor ? expectedColor : null;
            foreach (var marker in markers)
            {
                string? DecorationColor(string? css) => (css ?? "").Split(';').FirstOrDefault(p => p.StartsWith("text-decoration-color:", StringComparison.Ordinal))?.Split(':')[1];
                var css = marker.GetAttribute("style") ?? "";
                Assert.Equal(style == "words" ? null : color, DecorationColor(css));
                Assert.False(css.Split(';').Any(p => p.StartsWith("color:", StringComparison.Ordinal)));
                var words = marker.QuerySelectorAll(".dxp-marker-word-underline");
                if (style == "words")
                {
                    Assert.Equal(2, words.Length);
                    Assert.All(words, word => Assert.Equal(color, DecorationColor(word.GetAttribute("style"))));
                }
                else Assert.Empty(words);
                var separate = strikeLines == 1 && style is not ("words" or "none") &&
                    (color != null || style is "double" or "dotted" or "dash" or "wave");
                Assert.Equal(strikeLines == 1 && !separate, css.Contains("line-through", StringComparison.Ordinal));
                var singles = marker.QuerySelectorAll(".dxp-marker-single-strike");
                if (separate) Assert.Equal("text-decoration:line-through;text-decoration-style:solid", Assert.Single(singles).GetAttribute("style"));
                else Assert.Empty(singles);
                var doubles = marker.QuerySelectorAll(".dxp-marker-double-strike");
                if (strikeLines == 2) Assert.Equal("text-decoration:line-through;text-decoration-style:double", Assert.Single(doubles).GetAttribute("style"));
                else Assert.Empty(doubles);
            }
        }
        Assert.Equal(before, Snapshot());
    }
}
