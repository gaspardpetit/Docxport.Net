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

public sealed class DocJapaneseNumericMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "counting", "legal", "digits" })
            foreach (var pair in new[] { (0, 1), (0, 98), (0, 9998), (8, 9998) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void JapaneseNumericMarkersSurviveDoc(int depth, string mode)
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
            "counting" => (0x0B, NumberFormatValues.JapaneseCounting),
            "legal" => (0x10, NumberFormatValues.JapaneseLegal),
            "digits" => (0x11, NumberFormatValues.JapaneseDigitalTenThousand),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-15.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "counting-1" => ["BEGIN\u4e00END", "BEGIN\u4e8cEND", "BEGIN\u4e09END", "BEGIN\u56dbEND", "BEGIN\u4e94END", "BEGIN\u516dEND"],
        "counting-98" => ["BEGIN\u4e5d\u5341\u516bEND", "BEGIN\u4e5d\u5341\u4e5dEND", "BEGIN\u767eEND", "BEGIN\u767e\u4e00END", "BEGIN\u767e\u4e8cEND", "BEGIN\u767e\u4e09END"],
        "counting-9998" => ["BEGIN\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u516bEND", "BEGIN\u4e5d\u5343\u4e5d\u767e\u4e5d\u5341\u4e5dEND", "BEGIN\u4e00\u4e07END", "BEGIN\u4e00\u4e07\u4e00END", "BEGIN\u4e00\u4e07\u4e8cEND", "BEGIN\u4e00\u4e07\u4e09END"],
        "digits-1" => ["BEGIN\u4e00END", "BEGIN\u4e8cEND", "BEGIN\u4e09END", "BEGIN\u56dbEND", "BEGIN\u4e94END", "BEGIN\u516dEND"],
        "digits-98" => ["BEGIN\u4e5d\u516bEND", "BEGIN\u4e5d\u4e5dEND", "BEGIN\u4e00\u3007\u3007END", "BEGIN\u4e00\u3007\u4e00END", "BEGIN\u4e00\u3007\u4e8cEND", "BEGIN\u4e00\u3007\u4e09END"],
        "digits-9998" => ["BEGIN\u4e5d\u4e5d\u4e5d\u516bEND", "BEGIN\u4e5d\u4e5d\u4e5d\u4e5dEND", "BEGINEND", "BEGINEND", "BEGINEND", "BEGINEND"],
        "legal-1" => ["BEGIN\u58f1END", "BEGIN\u5f10END", "BEGIN\u53c2END", "BEGIN\u56dbEND", "BEGIN\u4f0dEND", "BEGIN\u516dEND"],
        "legal-98" => ["BEGIN\u4e5d\u62fe\u516bEND", "BEGIN\u4e5d\u62fe\u4e5dEND", "BEGIN\u58f1\u767eEND", "BEGIN\u58f1\u767e\u58f1END", "BEGIN\u58f1\u767e\u5f10END", "BEGIN\u58f1\u767e\u53c2END"],
        "legal-9998" => ["BEGIN\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u516bEND", "BEGIN\u4e5d\u9621\u4e5d\u767e\u4e5d\u62fe\u4e5dEND", "BEGIN\u58f1\u842cEND", "BEGIN\u58f1\u842c\u58f1END", "BEGIN\u58f1\u842c\u5f10END", "BEGIN\u58f1\u842c\u53c2END"],
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
