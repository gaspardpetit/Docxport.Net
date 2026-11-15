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

public sealed class DocArabicMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "alpha", "abjad" })
            foreach (var pair in new[] { (0, 1), (0, 27), (0, 55), (0, 391), (8, 391) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ArabicMarkersSurviveDoc(int depth, string mode)
    {
        var bytes = CreateSource(depth, mode);
        var binary = DxpDocExport.Export(bytes);
        using (var input = new MemoryStream(binary))
        using (var index = new DocTextIndexWalker().Index(input))
        {
            Assert.Single(index.Lists.Definitions);
            var instance = Assert.Single(index.Lists.Instances);
            Assert.Empty(instance.StartOverrides);
            Assert.Empty(instance.FormattingOverrides);
            Assert.Equal(Definition(mode).Code,
                index.Lists.Definitions.Single().Levels[depth].NumberFormat);
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
        Assert.Equal(1, projectedNumbering.Elements<AbstractNum>().Count());
        Assert.Single(projectedNumbering.Elements<NumberingInstance>());
        Assert.Empty(projectedNumbering.Descendants<LevelOverride>());
        Assert.Equal(Definition(mode).Format,
            projectedNumbering.Elements<AbstractNum>().Single().Elements<Level>()
                .Single(x => x.LevelIndex!.Value == depth).NumberingFormat!.Val!.Value);
        var expected = Labels(depth, mode);
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
                Assert.Equal(string.Concat(expected.Select((marker, i) => marker + "Item" + i)), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    private static (byte Code, NumberFormatValues Format) Definition(string mode) =>
        mode[..mode.LastIndexOf('-')] switch
        {
            "alpha" => (0x2E, NumberFormatValues.ArabicAlpha),
            "abjad" => (0x30, NumberFormatValues.ArabicAbjad),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-15.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "abjad-1" => ["BEGIN\u200c\u0623END", "BEGIN\u200c\u0628END", "BEGIN\u200c\u062cEND", "BEGIN\u200c\u062fEND", "BEGIN\u200c\u0647END", "BEGIN\u200c\u0648END"],
        "abjad-27" => ["BEGIN\u200c\u0638END", "BEGIN\u200c\u063aEND", "BEGIN\u200c\u0623\u200c\u0623END", "BEGIN\u200c\u0628\u200c\u0628END", "BEGIN\u200c\u062c\u200c\u062cEND", "BEGIN\u200c\u062f\u200c\u062fEND"],
        "abjad-391" => ["BEGIN\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638\u200c\u0638END", "BEGIN\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063a\u200c\u063aEND", "BEGIN\u200c\u0623END", "BEGIN\u200c\u0628END", "BEGIN\u200c\u062cEND", "BEGIN\u200c\u062fEND"],
        "abjad-55" => ["BEGIN\u200c\u0638\u200c\u0638END", "BEGIN\u200c\u063a\u200c\u063aEND", "BEGIN\u200c\u0623\u200c\u0623\u200c\u0623END", "BEGIN\u200c\u0628\u200c\u0628\u200c\u0628END", "BEGIN\u200c\u062c\u200c\u062c\u200c\u062cEND", "BEGIN\u200c\u062f\u200c\u062f\u200c\u062fEND"],
        "alpha-1" => ["BEGIN\u0623\u200cEND", "BEGIN\u0628\u200cEND", "BEGIN\u062a\u200cEND", "BEGIN\u062b\u200cEND", "BEGIN\u062c\u200cEND", "BEGIN\u062d\u200cEND"],
        "alpha-27" => ["BEGIN\u0648\u200cEND", "BEGIN\u064a\u200cEND", "BEGIN\u0623\u200c\u0623\u200cEND", "BEGIN\u0628\u200c\u0628\u200cEND", "BEGIN\u062a\u200c\u062a\u200cEND", "BEGIN\u062b\u200c\u062b\u200cEND"],
        "alpha-391" => ["BEGIN\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200c\u0648\u200cEND", "BEGIN\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200c\u064a\u200cEND", "BEGIN\u0623\u200cEND", "BEGIN\u0628\u200cEND", "BEGIN\u062a\u200cEND", "BEGIN\u062b\u200cEND"],
        "alpha-55" => ["BEGIN\u0648\u200c\u0648\u200cEND", "BEGIN\u064a\u200c\u064a\u200cEND", "BEGIN\u0623\u200c\u0623\u200c\u0623\u200cEND", "BEGIN\u0628\u200c\u0628\u200c\u0628\u200cEND", "BEGIN\u062a\u200c\u062a\u200c\u062a\u200cEND", "BEGIN\u062b\u200c\u062b\u200c\u062b\u200cEND"],
        _ => throw new ArgumentException("Unknown test case", nameof(mode))
    };

    public static byte[] CreateSource(int depth, string mode)
    {
        using var stream = new MemoryStream();
        stream.Write(DocOverrideFlagTests.CreateSource(depth, "none"));
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            var level = main.NumberingDefinitionsPart!.Numbering!.Elements<AbstractNum>().Single()
                .Elements<Level>().Single(x => x.LevelIndex!.Value == depth);
            level.StartNumberingValue!.Val = int.Parse(mode.Split('-').Last());
            level.NumberingFormat!.Val = Definition(mode).Format;
            level.AppendChild(new NumberingSymbolRunProperties(new RunFonts
                { Ascii = "Arial", HighAnsi = "Arial", EastAsia = "Arial" }));
        }
        return stream.ToArray();
    }
}
