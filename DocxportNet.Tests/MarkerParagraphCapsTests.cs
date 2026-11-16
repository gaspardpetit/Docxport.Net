using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxportNet.API;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class MarkerParagraphCapsTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var script in new[] { "latin", "greek" })
            foreach (var mark in new[] { "absent", "off", "on" })
                foreach (var level in new[] { "absent", "off", "on" })
                    yield return new object[] { script, mark, level };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ParagraphMarkCapsUsesLevelOverride(string script, string mark, string level)
    {
        var name = $"{script}-mark-{mark}-level-{level}";
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        Assert.NotNull(root);
        using var document = WordprocessingDocument.Open(Path.Combine(root.FullName,
            "fixtures/docx/regressions/list-markers/ParagraphMarkerCaps/source", name + ".docx"), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        var expected = script == "latin" ? "ab" : "\u03b1\u03b2";
        if (level == "on" || level == "absent" && mark == "on") expected = expected.ToUpperInvariant();
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
                var line = Assert.Single(text.Split('\n').Where(line => line.Contains("Item" + i)));
                Assert.True(line.TrimStart().StartsWith(expected, StringComparison.Ordinal), name + ": " + line);
            }
        }
        Assert.Equal(before, Snapshot());
    }
}
