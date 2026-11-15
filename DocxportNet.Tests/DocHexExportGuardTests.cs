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

public sealed class DocHexExportGuardTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var depth in new[] { 0, 8 })
            foreach (var mode in new[] { "definition", "override" })
                yield return new object[] { depth, mode };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void HexExportReportsUnsupportedInsteadOfWritingAnUnreadableDoc(int depth, string mode)
    {
        var source = CreateSource(depth, mode, "hex");
        Assert.Throws<NotSupportedException>(() => DxpDocExport.Export(source));
        CheckText(source, Labels("hex"));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void HistoricalHexDocsStillProjectAndRender(int depth, string mode)
    {
        var binary = File.ReadAllBytes(FixturePath.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "fixtures", "doc", "regressions", "list-markers",
            "HexExportGuard", $"hex-depth-{depth}-{mode}.doc"));
        using (var stream = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(stream))
        {
            var level = mode == "override" ? index.Lists.Instances.Single().FormattingOverrides[depth]
                : index.Lists.Definitions.Single().Levels[depth];
            Assert.Equal((byte)8, level.NumberFormat);
        }
        var projected = DxpDocToDocx.Project(binary).DocxBytes;
        using var input = new MemoryStream(projected);
        using var doc = WordprocessingDocument.Open(input, false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
        Assert.Contains(doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.Descendants<NumberingFormat>(),
            x => x.Val?.Value == NumberFormatValues.Hex);
        CheckText(projected, Labels("hex"));
        Assert.Throws<NotSupportedException>(() => DxpDocExport.Export(projected));
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void WordAcceptedHalfWidthDecimalRemainsWritable(int depth, string mode)
    {
        var source = CreateSource(depth, mode, "decimalHalfWidth");
        var binary = DxpDocExport.Export(source);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            var level = mode == "override" ? index.Lists.Instances.Single().FormattingOverrides[depth]
                : index.Lists.Definitions.Single().Levels[depth];
            Assert.Equal((byte)0x0f, level.NumberFormat);
        }
        CheckText(source, Labels("decimalHalfWidth"));
        var projected = DxpDocToDocx.Project(binary).DocxBytes;
        using var stream = new MemoryStream(projected);
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
        CheckText(projected, Labels("decimalHalfWidth"));
    }

    private static void CheckText(byte[] bytes, string[] labels)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        DxpITextVisitor[] visitors = [new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreatePlainConfig()),
            new DxpMarkdownVisitor(DxpMarkdownVisitorConfig.CreateRichConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreatePlainConfig()),
            new DxpHtmlVisitor(DxpHtmlVisitorConfig.CreateRichConfig()),
            new DxpPlainTextVisitor(DxpPlainTextVisitorConfig.CreateAcceptConfig())];
        foreach (var visitor in visitors)
        {
            var text = DxpExport.ExportToString(doc, visitor);
            if (visitor is DxpHtmlVisitor or DxpMarkdownVisitor)
                text = new HtmlParser().ParseDocument(text).Body!.TextContent;
            Assert.Equal(string.Concat(labels.Select((label, i) => label + "Item" + i)), Regex.Replace(text, @"\s+", ""));
        }
    }

    // Source and projected editable ListString values captured from Word on 2026-11-15.
    public static string[] Labels(string format) => format == "hex"
        ? ["BEGINEEND", "BEGINFEND", "BEGIN10END", "BEGIN11END", "BEGIN12END", "BEGIN13END"]
        : ["BEGIN14END", "BEGIN15END", "BEGIN16END", "BEGIN17END", "BEGIN18END", "BEGIN19END"];

    public static byte[] CreateSource(int depth, string mode, string format)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, mode == "override" ? "format" : "none"));
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var numbering = doc.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
            // Formatting-only overrides inherit the abstract start in Word.
            numbering.Elements<AbstractNum>().Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth)
                .StartNumberingValue!.Val = 14;
            var level = mode == "override" ? numbering.Descendants<LevelOverride>().Single().Level!
                : numbering.Elements<AbstractNum>().Single().Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
            level.StartNumberingValue!.Val = 14;
            level.NumberingFormat!.Val = new NumberFormatValues(format);
            level.AppendChild(new NumberingSymbolRunProperties(new RunFonts { Ascii = "Arial", HighAnsi = "Arial" }));
        }
        return stream.ToArray();
    }
}
