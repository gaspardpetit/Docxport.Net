using System.Text.RegularExpressions;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxportNet.API;
using DocxportNet.Doc;
using DocxportNet.Visitors.Html;
using DocxportNet.Visitors.Markdown;
using DocxportNet.Visitors.PlainText;

namespace DocxportNet.Tests;

public sealed class DocMarkerComplexFontSlotTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var mode in new[] { "definition-only", "override-only", "definition-mixed", "override-mixed" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ComplexScriptOnlyMarkerFontSurvivesNativeReadAndProjection(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        // Keep exercising the original CS-only binary slots: new exports materialize
        // a forced CS face into all ordinary slots for Word's DOC marker path.
        var originalBinary = File.ReadAllBytes(FixturePath.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "fixtures", "doc", "regressions", "list-markers",
            "MarkerComplexFontSlot", $"cs-font-depth-{depth}-{mode}.doc"));
        using (var input = new MemoryStream(originalBinary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var level = mode.StartsWith("override", StringComparison.Ordinal)
                ? index.Lists.Instances.Single().FormattingOverrides[depth]
                : index.Lists.Definitions.Single().Levels[depth];
            Assert.NotNull(level.LabelFormatting!.ComplexScriptFontIndex);
            Assert.Equal("Courier New", level.LabelFormatting.ComplexScriptFontName);
            Assert.Equal(mode.EndsWith("mixed", StringComparison.Ordinal) ? "Arial" : null, level.LabelFormatting.AsciiFontName);
        }
        using var sourceStream = new MemoryStream(bytes);
        using var source = WordprocessingDocument.Open(sourceStream, false);
        var numbering = source.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var original = numbering.OuterXml;
        _ = DocOpenXmlListReader.Read(source.MainDocumentPart, _ => DocCharacterFormatting.Empty, _ => DocParagraphFormatting.Empty);
        Assert.Equal(original, numbering.OuterXml);
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(binary).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        Assert.Empty(new OpenXmlValidator().Validate(projected));
        var projectedNumbering = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var target = mode.StartsWith("override", StringComparison.Ordinal)
            ? projectedNumbering.Descendants<LevelOverride>().Single().Level!
            : projectedNumbering.Elements<AbstractNum>().Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
        var props = target.NumberingSymbolRunProperties!;
        Assert.Equal("Courier New", props.RunFonts?.ComplexScript?.Value);
        Assert.NotNull(props.GetFirstChild<ComplexScript>());
        foreach (var document in new[] { source, projected })
        {
            DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
                new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
                new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
                new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
            foreach (var visitor in visitors)
            {
                var output = DxpExport.ExportToString(document, visitor);
                if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor)
                    output = new HtmlParser().ParseDocument(output).Body!.TextContent;
                Assert.Equal(string.Concat(Labels(depth, mode).Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    public static string[] Labels(int depth, string mode) => Enumerable.Repeat("\u05d0", 6).ToArray();

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode.StartsWith("override", StringComparison.Ordinal) ? "format" : "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            var level = mode.StartsWith("override", StringComparison.Ordinal)
                ? numbering.Descendants<LevelOverride>().Single().Level!
                : numbering.Elements<AbstractNum>().Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
            level.NumberingFormat!.Val = NumberFormatValues.Bullet;
            level.LevelText!.Val = "\u05d0";
            var fonts = new RunFonts { ComplexScript = "Courier New" };
            if (mode.EndsWith("mixed", StringComparison.Ordinal))
            {
                fonts.Ascii = "Arial"; fonts.HighAnsi = "Times New Roman"; fonts.EastAsia = "MS Mincho";
            }
            level.AppendChild(new NumberingSymbolRunProperties(fonts, new ComplexScript()));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts
                    { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }, new FontSize { Val = "24" }))));
        }
        return stream.ToArray();
    }
}
