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

public sealed class DocHebrewMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "numeric", "alphabetic" })
            foreach (var pair in new[] { (0, 1), (0, 14), (0, 21), (0, 391), (8, 391) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void HebrewMarkersSurviveDoc(int depth, string mode)
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
            "numeric" => (0x2D, NumberFormatValues.Hebrew1),
            "alphabetic" => (0x2F, NumberFormatValues.Hebrew2),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-14.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "alphabetic-1" => ["BEGIN\u200f\u05d0END", "BEGIN\u200f\u05d1END", "BEGIN\u200f\u05d2END", "BEGIN\u200f\u05d3END", "BEGIN\u200f\u05d4END", "BEGIN\u200f\u05d5END"],
        "alphabetic-14" => ["BEGIN\u200f\u05e0END", "BEGIN\u200f\u05e1END", "BEGIN\u200f\u05e2END", "BEGIN\u200f\u05e4END", "BEGIN\u200f\u05e6END", "BEGIN\u200f\u05e7END"],
        "alphabetic-21" => ["BEGIN\u200f\u05e9END", "BEGIN\u200f\u05eaEND", "BEGIN\u200f\u05ea\u05d0END", "BEGIN\u200f\u05ea\u05d1END", "BEGIN\u200f\u05ea\u05d2END", "BEGIN\u200f\u05ea\u05d3END"],
        "alphabetic-391" => ["BEGIN\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e4END", "BEGIN\u200f\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05ea\u05e6END", "BEGIN\u200f\u05d0END", "BEGIN\u200f\u05d1END", "BEGIN\u200f\u05d2END", "BEGIN\u200f\u05d3END"],
        "numeric-1" => ["BEGIN\u05d0END", "BEGIN\u05d1END", "BEGIN\u05d2END", "BEGIN\u05d3END", "BEGIN\u05d4END", "BEGIN\u05d5END"],
        "numeric-14" => ["BEGIN\u05d9\u05d3END", "BEGIN\u05d8\u05d5END", "BEGIN\u05d8\u05d6END", "BEGIN\u05d9\u05d6END", "BEGIN\u05d9\u05d7END", "BEGIN\u05d9\u05d8END"],
        "numeric-21" => ["BEGIN\u05db\u05d0END", "BEGIN\u05db\u05d1END", "BEGIN\u05db\u05d2END", "BEGIN\u05db\u05d3END", "BEGIN\u05db\u05d4END", "BEGIN\u05db\u05d5END"],
        "numeric-391" => ["BEGIN\u05e9\u05e6\u05d0END", "BEGIN\u05e9\u05e6\u05d1END", "BEGIN\u05d0END", "BEGIN\u05d1END", "BEGIN\u05d2END", "BEGIN\u05d3END"],
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
