using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class MarkerStyleCapsTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/StyleMarkerCaps");
    }

    public static IEnumerable<object[]> Cases()
    {
        using var json = System.Text.Json.JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            yield return new object[] { row.GetProperty("case").GetString()!, row.GetProperty("expectedVisibleLabels")[0].GetString()! };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void InheritsStyleCapsWithoutChangingSource(string name, string expected)
    {
        using var document = WordprocessingDocument.Open(Path.Combine(FixtureRoot(), "source", name + ".docx"), false);
        using (var validationDocument = WordprocessingDocument.Open(Path.Combine(FixtureRoot(), "source", name + ".docx"), false,
            new OpenSettings { MarkupCompatibilityProcessSettings = new(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2013) }))
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(validationDocument));
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        foreach (var visitor in new DxpITextVisitor[] {
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig()) })
        {
            var output = DxpExport.ExportToString(document, visitor);
            var text = visitor is DxpPlainTextVisitor ? output : new HtmlParser().ParseDocument(output).Body!.TextContent;
            for (var i = 0; i < 3; i++)
            {
                var line = Assert.Single(text.Split('\n'), line => line.Contains("Item" + i));
                Assert.True(line.TrimStart().StartsWith(expected, StringComparison.Ordinal), name + ": " + line);
            }
        }
        Assert.Equal(before, Snapshot());
    }
}
