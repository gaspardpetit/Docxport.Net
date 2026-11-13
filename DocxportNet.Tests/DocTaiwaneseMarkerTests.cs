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

public sealed class DocTaiwaneseMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "counting", "legal", "thousand", "digital" })
            foreach (var pair in new[] { (0, 1), (0, 98), (0, 9998), (8, 9998) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void TaiwaneseMarkersSurviveDoc(int depth, string mode)
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
            "counting" => (0x21, NumberFormatValues.TaiwaneseCounting),
            "legal" => (0x22, NumberFormatValues.IdeographLegalTraditional),
            "thousand" => (0x23, NumberFormatValues.TaiwaneseCountingThousand),
            "digital" => (0x24, NumberFormatValues.TaiwaneseDigital),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-13.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "counting-1" => ["BEGIN一END", "BEGIN二END", "BEGIN三END", "BEGIN四END", "BEGIN五END", "BEGIN六END"],
        "counting-98" => ["BEGIN九十八END", "BEGIN九十九END", "BEGIN一○○END", "BEGIN一○一END", "BEGIN一○二END", "BEGIN一○三END"],
        "counting-9998" => ["BEGIN九九九八END", "BEGIN九九九九END", "BEGIN一○○○○END", "BEGIN一○○○一END", "BEGIN一○○○二END", "BEGIN一○○○三END"],
        "digital-1" => ["BEGIN一END", "BEGIN二END", "BEGIN三END", "BEGIN四END", "BEGIN五END", "BEGIN六END"],
        "digital-98" => ["BEGIN九八END", "BEGIN九九END", "BEGIN一○○END", "BEGIN一○一END", "BEGIN一○二END", "BEGIN一○三END"],
        "digital-9998" => ["BEGIN九九九八END", "BEGIN九九九九END", "BEGIN一○○○○END", "BEGIN一○○○一END", "BEGIN一○○○二END", "BEGIN一○○○三END"],
        "legal-1" => ["BEGIN壹END", "BEGIN貳END", "BEGIN參END", "BEGIN肆END", "BEGIN伍END", "BEGIN陸END"],
        "legal-98" => ["BEGIN玖拾捌END", "BEGIN玖拾玖END", "BEGIN壹佰END", "BEGIN壹佰零壹END", "BEGIN壹佰零貳END", "BEGIN壹佰零參END"],
        "legal-9998" => ["BEGIN玖仟玖佰玖拾捌END", "BEGIN玖仟玖佰玖拾玖END", "BEGIN壹萬END", "BEGIN壹萬零壹END", "BEGIN壹萬零貳END", "BEGIN壹萬零參END"],
        "thousand-1" => ["BEGIN一END", "BEGIN二END", "BEGIN三END", "BEGIN四END", "BEGIN五END", "BEGIN六END"],
        "thousand-98" => ["BEGIN九十八END", "BEGIN九十九END", "BEGIN一百END", "BEGIN一百零一END", "BEGIN一百零二END", "BEGIN一百零三END"],
        "thousand-9998" => ["BEGIN九千九百九十八END", "BEGIN九千九百九十九END", "BEGIN一萬END", "BEGIN一萬零一END", "BEGIN一萬零二END", "BEGIN一萬零三END"],
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
                { Ascii = "Yu Gothic", HighAnsi = "Yu Gothic", EastAsia = "Yu Gothic" }));
        }
        return stream.ToArray();
    }
}
