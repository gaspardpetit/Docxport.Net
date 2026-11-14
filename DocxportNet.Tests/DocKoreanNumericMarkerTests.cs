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

public sealed class DocKoreanNumericMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "digital", "counting", "legal", "digital2" })
            foreach (var pair in new[] { (0, 1), (0, 98), (0, 9998), (8, 9998) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void KoreanNumericMarkersSurviveDoc(int depth, string mode)
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
            "digital" => (0x29, NumberFormatValues.KoreanDigital),
            "counting" => (0x2A, NumberFormatValues.KoreanCounting),
            "legal" => (0x2B, NumberFormatValues.KoreanLegal),
            "digital2" => (0x2C, NumberFormatValues.KoreanDigital2),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-13.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "counting-1" => ["BEGIN일END", "BEGIN이END", "BEGIN삼END", "BEGIN사END", "BEGIN오END", "BEGIN육END"],
        "counting-98" => ["BEGIN구십팔END", "BEGIN구십구END", "BEGIN백END", "BEGIN백일END", "BEGIN백이END", "BEGIN백삼END"],
        "counting-9998" => ["BEGIN구천구백구십팔END", "BEGIN구천구백구십구END", "BEGIN만END", "BEGIN만일END", "BEGIN만이END", "BEGIN만삼END"],
        "digital-1" => ["BEGIN일END", "BEGIN이END", "BEGIN삼END", "BEGIN사END", "BEGIN오END", "BEGIN육END"],
        "digital-98" => ["BEGIN구팔END", "BEGIN구구END", "BEGIN일영영END", "BEGIN일영일END", "BEGIN일영이END", "BEGIN일영삼END"],
        "digital-9998" => ["BEGIN구구구팔END", "BEGIN구구구구END", "BEGIN일영영영영END", "BEGIN일영영영일END", "BEGIN일영영영이END", "BEGIN일영영영삼END"],
        "digital2-1" => ["BEGIN一END", "BEGIN二END", "BEGIN三END", "BEGIN四END", "BEGIN五END", "BEGIN六END"],
        "digital2-98" => ["BEGIN九八END", "BEGIN九九END", "BEGIN一零零END", "BEGIN一零一END", "BEGIN一零二END", "BEGIN一零三END"],
        "digital2-9998" => ["BEGIN九九九八END", "BEGIN九九九九END", "BEGIN一零零零零END", "BEGIN一零零零一END", "BEGIN一零零零二END", "BEGIN一零零零三END"],
        "legal-1" => ["BEGIN하나END", "BEGIN둘END", "BEGIN셋END", "BEGIN넷END", "BEGIN다섯END", "BEGIN여섯END"],
        "legal-98" => ["BEGIN아흔여덟END", "BEGIN아흔아홉END", "BEGIN백END", "BEGIN백일END", "BEGIN백이END", "BEGIN백삼END"],
        "legal-9998" => ["BEGIN구천구백구십팔END", "BEGIN구천구백구십구END", "BEGIN만END", "BEGIN만일END", "BEGIN만이END", "BEGIN만삼END"],
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
                { Ascii = "Malgun Gothic", HighAnsi = "Malgun Gothic", EastAsia = "Malgun Gothic" }));
        }
        return stream.ToArray();
    }
}
