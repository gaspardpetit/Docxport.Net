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

public sealed class DocMarkerForcedComplexFontTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var mode in new[] { "definition-only", "override-only", "definition-mixed", "override-mixed", "definition-disabled", "override-disabled" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ForcedComplexMarkerFontSurvivesDocExport(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var level = mode.StartsWith("override", StringComparison.Ordinal)
                ? index.Lists.Instances.Single().FormattingOverrides[depth]
                : index.Lists.Definitions.Single().Levels[depth];
            Assert.NotNull(level.LabelFormatting!.ComplexScriptFontIndex);
            Assert.Equal("Courier New", level.LabelFormatting.ComplexScriptFontName);
            Assert.Equal(mode.EndsWith("disabled", StringComparison.Ordinal) ? "Arial" : "Courier New", level.LabelFormatting.AsciiFontName);
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
        Assert.Equal(!mode.EndsWith("disabled", StringComparison.Ordinal), props.GetFirstChild<ComplexScript>()?.Val?.Value ?? true);
        Assert.Equal(mode.EndsWith("disabled", StringComparison.Ordinal) ? "Arial" : "Courier New", props.RunFonts?.Ascii?.Value);
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SymbolComplexFontKeepsItsEncodingSlots(bool force)
    {
        using var stream = new MemoryStream();
        stream.Write(CreateSource(0, "definition-mixed"));
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var props = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Descendants<NumberingSymbolRunProperties>().Single();
            props.RunFonts!.ComplexScript = "Symbol";
            props.GetFirstChild<ComplexScript>()!.Val = force;
        }
        using var projectedStream = new MemoryStream(DxpDocToDocx.Project(DxpDocExport.Export(stream.ToArray())).DocxBytes);
        using var projected = WordprocessingDocument.Open(projectedStream, false);
        var fonts = projected.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Descendants<NumberingSymbolRunProperties>().Single().RunFonts!;
        Assert.Equal("Arial", fonts.Ascii?.Value);
        Assert.Equal("Times New Roman", fonts.HighAnsi?.Value);
        Assert.Equal("MS Mincho", fonts.EastAsia?.Value);
        Assert.Equal("Symbol", fonts.ComplexScript?.Value);
    }

    public static string[] Labels(int depth, string mode) => Enumerable.Repeat("BEGINabEND", 6).ToArray();

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
            level.LevelText!.Val = "BEGINabEND";
            var fonts = new RunFonts { ComplexScript = "Courier New" };
            if (!mode.EndsWith("only", StringComparison.Ordinal))
            {
                fonts.Ascii = "Arial"; fonts.HighAnsi = "Times New Roman"; fonts.EastAsia = "MS Mincho";
            }
            level.AppendChild(new NumberingSymbolRunProperties(fonts, new ComplexScript { Val = !mode.EndsWith("disabled", StringComparison.Ordinal) }));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts
                    { Ascii = "Arial", HighAnsi = "Arial", ComplexScript = "Arial" }, new FontSize { Val = "24" }))));
        }
        return stream.ToArray();
    }
}
