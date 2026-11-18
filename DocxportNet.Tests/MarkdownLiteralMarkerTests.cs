using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.Markdown;
using System.Text.Json;

namespace DocxportNet.Tests;

public sealed class MarkdownLiteralMarkerTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlLiteralMarkerMarkup");
    }

    public static IEnumerable<object[]> Cases()
    {
        foreach (var row in HtmlLiteralMarkerMarkupTests.Cases())
            foreach (var rich in new[] { true, false })
                yield return [row[0], "source/" + row[0] + ".docx", row[1], false, rich];
        using var json = JsonDocument.Parse(File.ReadAllText(Path.Combine(FixtureRoot(), "markdown-word-evidence.json")));
        foreach (var row in json.RootElement.GetProperty("cases").EnumerateArray())
            foreach (var rich in new[] { true, false })
                yield return [row.GetProperty("case").GetString()!, "markdown/source/" + row.GetProperty("case").GetString() + ".docx",
                    row.GetProperty("expectedVisibleLabels").EnumerateArray().Select(label => label.GetString()!).ToArray(),
                    row.GetProperty("context").GetString() == "standalone", rich];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void PreservesLiteralHtmlCharacters(string name, string source, string[] labels, bool standalone, bool rich)
    {
        using var document = WordprocessingDocument.Open(Path.Combine(FixtureRoot(), source), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        var output = Export(document, rich);
        // These labels exercise HTML characters, not Markdown punctuation.
        // External CLI evidence also renders the Markdown with markdown-it.
        var html = new HtmlParser().ParseDocument(output);
        Assert.Empty(html.QuerySelectorAll("b,z,span:not([style])"));
        var lines = html.Body!.TextContent.Split('\n').Select(line => line.Trim()).Where(line => line.Length != 0).ToArray();
        if (standalone) Assert.Equal(labels, lines);
        else
            for (var i = 0; i < labels.Length; i++)
            {
                var line = Assert.Single(lines, line => line.Contains("Item" + i));
                Assert.Equal(labels[i] + " Item" + i, line);
            }
        Assert.Equal(before, Snapshot());
    }

    public static IEnumerable<object[]> OrderedCases() =>
        from delimiter in new[] { ".", ")" }
        from context in new[] { "ordinary", "wrapped", "standalone" }
        from rich in new[] { true, false }
        select new object[] { delimiter, context, rich };

    [Theory]
    [MemberData(nameof(OrderedCases))]
    public void PreservesOrderedListEscaping(string delimiter, string context, bool rich)
    {
        using var stream = LoadSource("plain-numbered");
        using var document = WordprocessingDocument.Open(stream, true);
        document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Descendants<Level>().Single().LevelText!.Val = "%1" + delimiter;
        foreach (var paragraph in document.MainDocumentPart.Document!.Descendants<Paragraph>())
            if (context == "standalone") paragraph.RemoveAllChildren<Run>();
            else if (context == "wrapped") paragraph.ParagraphProperties!.SpacingBetweenLines = new() { Before = "120", After = "120" };
        var output = Export(document, rich).Replace("\r\n", "\n");
        if (context == "wrapped" && rich) Assert.Contains(">1" + delimiter + " Item0", output);
        else Assert.Contains("1" + (delimiter == "." ? "&#46;" : "&#41;") +
            (context == "standalone" ? "\n" : " Item0"), output);
    }

    [Theory]
    [InlineData("Symbol", "\u00b7", "•")]
    [InlineData("Wingdings", "A", "✌")]
    public void PreservesSymbolFontTranslation(string font, string character, string expected)
    {
        using var stream = LoadSource("plain-font-bullet");
        using var document = WordprocessingDocument.Open(stream, true);
        var level = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Descendants<Level>().Single();
        level.LevelText!.Val = character;
        level.NumberingSymbolRunProperties!.RunFonts!.Ascii = font;
        level.NumberingSymbolRunProperties.RunFonts.HighAnsi = font;
        Assert.Contains($"<span style=\"font-family:{font}\">{character}</span>", Export(document, true));
        Assert.Contains(expected + " Item0", Export(document, false));
    }

    private static MemoryStream LoadSource(string name)
    {
        var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(Path.Combine(FixtureRoot(), "source", name + ".docx")));
        stream.Position = 0;
        return stream;
    }

    private static string Export(WordprocessingDocument document, bool rich) =>
        DxpExport.ExportToString(document, new DxpMarkdownVisitor(rich
            ? DxpMarkdownVisitorConfig.CreateRichConfig() : DxpMarkdownVisitorConfig.CreatePlainConfig()));
}
