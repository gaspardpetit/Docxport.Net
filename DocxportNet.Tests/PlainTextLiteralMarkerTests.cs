using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class PlainTextLiteralMarkerTests
{
    private static string FixtureRoot()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root != null && !Directory.Exists(Path.Combine(root.FullName, "fixtures"))) root = root.Parent;
        return Path.Combine(root!.FullName, "fixtures/docx/regressions/list-markers/HtmlLiteralMarkerMarkup");
    }

    public static IEnumerable<object[]> Cases()
    {
        // Reuse the exact generic sources and Word-visible labels from the HTML fix.
        foreach (var row in HtmlLiteralMarkerMarkupTests.Cases())
            foreach (var reject in new[] { false, true })
                yield return [row[0], row[1], reject];
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void PreservesLiteralMarkerText(string name, string[] labels, bool reject)
    {
        using var document = WordprocessingDocument.Open(Path.Combine(FixtureRoot(), "source", name + ".docx"), false);
        Verify(document, labels, reject);
    }

    [Theory]
    [InlineData("Symbol", "\u00b7", "•")]
    [InlineData("Wingdings", "A", "✌")]
    public void PreservesEstablishedSymbolFontTranslation(string font, string character, string expected)
    {
        using var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(Path.Combine(FixtureRoot(), "source", "plain-font-bullet.docx")));
        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, true);
        var level = document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Descendants<Level>().Single();
        level.LevelText!.Val = character;
        level.NumberingSymbolRunProperties!.RunFonts!.Ascii = font;
        level.NumberingSymbolRunProperties.RunFonts.HighAnsi = font;
        foreach (var reject in new[] { false, true })
            Verify(document, [expected, expected, expected], reject);
    }

    private static void Verify(WordprocessingDocument document, string[] labels, bool reject)
    {
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2013).Validate(document));
        var main = document.MainDocumentPart!;
        _ = main.Document!.Descendants().ToArray();
        _ = main.StyleDefinitionsPart!.Styles!.Descendants().ToArray();
        _ = main.NumberingDefinitionsPart!.Numbering!.Descendants().ToArray();
        string[] Snapshot() => [main.Document.OuterXml, main.StyleDefinitionsPart.Styles.OuterXml,
            main.NumberingDefinitionsPart.Numbering.OuterXml];
        var before = Snapshot();
        var text = DxpExport.ExportToString(document, new DxpPlainTextVisitor(reject
            ? DxpPlainTextVisitorConfig.CreateRejectConfig() : DxpPlainTextVisitorConfig.CreateAcceptConfig()));
        for (var i = 0; i < labels.Length; i++)
        {
            var line = Assert.Single(text.Split('\n'), line => line.Contains("Item" + i));
            Assert.Equal(labels[i] + " Item" + i, line.Trim());
        }
        Assert.Equal(before, Snapshot());
    }
}
