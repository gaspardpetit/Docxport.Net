using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class MarkdownOrderedMarkerEscapingTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/MarkdownOrderedMarkerEscaping");
    }

    public static IEnumerable<object[]> Cases()
    {
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            foreach (var mode in new[] { "rich", "plain", "no-colors", "no-layout" })
                yield return [row.GetProperty("case").GetString()!, row.GetProperty("context").GetString()!,
                    row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(label => label.GetString()!).ToArray(), mode];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void UsesAnEscapeValidInMarkdownAndHtml(string name, string context, string[] labels, string mode)
    {
        using var document = WordprocessingDocument.Open(Path.Combine(FixtureRoot(), "source", name + ".docx"), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        var main = document.MainDocumentPart!;
        foreach (var part in main.Parts) _ = part.OpenXmlPart.RootElement?.Descendants().ToArray();
        _ = main.Document!.Descendants().ToArray();
        string[] Snapshot() => new[] { main.Document.OuterXml }
            .Concat(main.Parts.OrderBy(part => part.RelationshipId).Select(part => part.OpenXmlPart.RootElement?.OuterXml ?? "")).ToArray();
        var before = Snapshot();
        var config = mode switch {
            "plain" => DxpMarkdownVisitorConfig.CreatePlainConfig(),
            "no-colors" => DxpMarkdownVisitorConfig.CreateRichConfig() with { EmitDocumentColors = false },
            "no-layout" => DxpMarkdownVisitorConfig.CreateRichConfig() with { EmitRichLayoutHtml = false },
            _ => DxpMarkdownVisitorConfig.CreateRichConfig()
        };
        var output = DxpExport.ExportToString(document, new DxpMarkdownVisitor(config));
        // A backslash escape is visible in raw HTML. The representation must
        // preserve the same literal label in both syntactic contexts.
        Assert.DoesNotContain("\\.", output);
        Assert.DoesNotContain("\\)", output);
        var html = new HtmlParser().ParseDocument(output);
        Assert.Empty(html.QuerySelectorAll("ol"));
        var lines = html.Body!.TextContent.Split('\n').Select(line => line.Trim()).Where(line => line.Length != 0).ToArray();
        if (context == "standalone") Assert.Equal(labels, lines);
        else
            for (var i = 0; i < labels.Length; i++)
            {
                var line = Assert.Single(lines, line => line.Contains("Item" + i));
                Assert.Equal(labels[i] + " Item" + i, line);
            }
        Assert.Equal(before, Snapshot());
    }
}
