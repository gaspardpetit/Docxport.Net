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

public sealed class DocMarkerFontHintTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var mode in new[] { "definition", "override" })
                foreach (var hint in new[] { "default", "eastAsia", "cs", "absent" })
                    foreach (var fontNames in new[] { false, true })
                        yield return new object[] { depth, mode, hint, fontNames };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void DirectFontHintSurvivesDocRoundTrip(int depth, string mode, string hint, bool fontNames)
    {
        var bytes = CreateSource(depth, mode, hint, fontNames);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var level = mode == "override" ? index.Lists.Instances.Single().FormattingOverrides[depth]
                : index.Lists.Definitions.Single().Levels[depth];
            byte? expected = hint switch { "default" => 0, "eastAsia" => 1, "cs" => 2, _ => null };
            Assert.Equal(expected, level.LabelFormatting?.FontHint);
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
        Assert.Equal(hint == "absent" ? null : hint, props?.RunFonts?.Hint?.InnerText);
        if (fontNames)
        {
            Assert.Equal("Arial", props?.RunFonts?.Ascii?.Value);
            Assert.Equal("Times New Roman", props?.RunFonts?.HighAnsi?.Value);
            Assert.Equal("MS Mincho", props?.RunFonts?.EastAsia?.Value);
            Assert.Equal("Courier New", props?.RunFonts?.ComplexScript?.Value);
        }
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
                Assert.Equal(string.Concat(Labels(depth, mode, hint, fontNames).Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    [Theory]
    [InlineData((byte)0)]
    [InlineData((byte)1)]
    [InlineData((byte)2)]
    [InlineData((byte)255)]
    public void NativeHintCodecRetainsAllDefinedOperands(byte hint)
    {
        var props = DocCharacterFormattingReader.ParseSprms(new byte[] { 0x6F, 0x28, hint });
        Assert.Equal(hint, props.FontHint);
        Assert.False(props.IsEmpty);
        Assert.Equal(new byte[] { 0x6F, 0x28, hint }, DocPlainTextWriter.EncodeCharacterProperties(props));
    }

    [Theory]
    [InlineData((byte)3)]
    [InlineData((byte)254)]
    public void NativeHintCodecRejectsUndefinedOperands(byte hint)
    {
        Assert.Throws<InvalidDataException>(() => DocCharacterFormattingReader.ParseSprms(new byte[] { 0x6F, 0x28, hint }));
        Assert.Throws<InvalidDataException>(() => DocPlainTextWriter.EncodeCharacterProperties(DocCharacterFormatting.Empty with { FontHint = hint }));
    }

    [Fact]
    public void PlainResetClearsTheDirectFontHint()
    {
        Assert.Null(DocCharacterFormattingReader.ParseSprms(new byte[] { 0x6F, 0x28, 1, 0x33, 0x2A, 0 }).FontHint);
    }

    public static string[] Labels(int depth, string mode, string hint, bool fontNames) => Enumerable.Repeat("\u03b1\u03b2", 3).ToArray();

    public static byte[] CreateSource(int depth, string mode, string hint, bool fontNames)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode.StartsWith("override", StringComparison.Ordinal) ? "format" : "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            foreach (var paragraph in main.Document.Body!.Elements<Paragraph>().Skip(3).ToArray()) paragraph.Remove();
            var numbering = main.NumberingDefinitionsPart!.Numbering!;
            var level = mode.StartsWith("override", StringComparison.Ordinal)
                ? numbering.Descendants<LevelOverride>().Single().Level!
                : numbering.Elements<AbstractNum>().Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
            level.NumberingFormat!.Val = NumberFormatValues.Bullet;
            level.LevelText!.Val = "\u03b1\u03b2";
            var fonts = new RunFonts();
            if (hint != "absent") fonts.Hint = new FontTypeHintValues(hint);
            if (fontNames)
            {
                fonts.ComplexScript = "Courier New"; fonts.Ascii = "Arial"; fonts.HighAnsi = "Times New Roman"; fonts.EastAsia = "MS Mincho";
            }
            level.AppendChild(new NumberingSymbolRunProperties(fonts));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts
                    { Ascii = "Arial", HighAnsi = "Times New Roman", EastAsia = "MS Mincho", ComplexScript = "Courier New" }, new FontSize { Val = "24" }))));
        }
        return stream.ToArray();
    }
}
