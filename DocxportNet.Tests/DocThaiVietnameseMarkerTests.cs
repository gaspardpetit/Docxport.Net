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

public sealed class DocThaiVietnameseMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "letters", "digits", "thai", "vietnamese" })
        {
            var boundaryStart = format == "letters" ? 40 : 98;
            var highStart = format == "letters" ? 1229 : format == "vietnamese" ? 998 : 9998;
            foreach (var pair in new[] { (0, 1), (0, boundaryStart), (0, highStart), (8, highStart) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
        }
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void ThaiVietnameseMarkersSurviveDoc(int depth, string mode)
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
                Assert.Equal(Regex.Replace(string.Concat(expected.Select((marker, i) => marker + "Item" + i)), @"\s+", ""), Regex.Replace(output, @"\s+", ""));
            }
        }
    }

    private static (byte Code, NumberFormatValues Format) Definition(string mode) =>
        mode[..mode.LastIndexOf('-')] switch
        {
            "letters" => (0x35, NumberFormatValues.ThaiLetters),
            "digits" => (0x36, NumberFormatValues.ThaiNumbers),
            "thai" => (0x37, NumberFormatValues.ThaiCounting),
            "vietnamese" => (0x38, NumberFormatValues.VietnameseCounting),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-15.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "digits-1" => ["BEGIN\u0e51END", "BEGIN\u0e52END", "BEGIN\u0e53END", "BEGIN\u0e54END", "BEGIN\u0e55END", "BEGIN\u0e56END"],
        "digits-98" => ["BEGIN\u0e59\u0e58END", "BEGIN\u0e59\u0e59END", "BEGIN\u0e51\u0e50\u0e50END", "BEGIN\u0e51\u0e50\u0e51END", "BEGIN\u0e51\u0e50\u0e52END", "BEGIN\u0e51\u0e50\u0e53END"],
        "digits-9998" => ["BEGIN\u0e59\u0e59\u0e59\u0e58END", "BEGIN\u0e59\u0e59\u0e59\u0e59END", "BEGIN\u0e51\u0e50\u0e50\u0e50\u0e50END", "BEGIN\u0e51\u0e50\u0e50\u0e50\u0e51END", "BEGIN\u0e51\u0e50\u0e50\u0e50\u0e52END", "BEGIN\u0e51\u0e50\u0e50\u0e50\u0e53END"],
        "letters-1" => ["BEGIN\u0e01END", "BEGIN\u0e02END", "BEGIN\u0e04END", "BEGIN\u0e07END", "BEGIN\u0e08END", "BEGIN\u0e09END"],
        "letters-1229" => ["BEGIN\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2d\u0e2dEND", "BEGIN\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2e\u0e2eEND", "BEGIN\u0e01END", "BEGIN\u0e02END", "BEGIN\u0e04END", "BEGIN\u0e07END"],
        "letters-40" => ["BEGIN\u0e2dEND", "BEGIN\u0e2eEND", "BEGIN\u0e01\u0e01END", "BEGIN\u0e02\u0e02END", "BEGIN\u0e04\u0e04END", "BEGIN\u0e07\u0e07END"],
        "thai-1" => ["BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07END", "BEGIN\u0e2a\u0e2d\u0e07END", "BEGIN\u0e2a\u0e32\u0e21END", "BEGIN\u0e2a\u0e35\u0e48END", "BEGIN\u0e2b\u0e49\u0e32END", "BEGIN\u0e2b\u0e01END"],
        "thai-98" => ["BEGIN\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e41\u0e1b\u0e14END", "BEGIN\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e40\u0e01\u0e49\u0e32END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e23\u0e49\u0e2d\u0e22END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e23\u0e49\u0e2d\u0e22\u0e40\u0e2d\u0e47\u0e14END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e23\u0e49\u0e2d\u0e22\u0e2a\u0e2d\u0e07END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e23\u0e49\u0e2d\u0e22\u0e2a\u0e32\u0e21END"],
        "thai-9998" => ["BEGIN\u0e40\u0e01\u0e49\u0e32\u0e1e\u0e31\u0e19\u0e40\u0e01\u0e49\u0e32\u0e23\u0e49\u0e2d\u0e22\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e41\u0e1b\u0e14END", "BEGIN\u0e40\u0e01\u0e49\u0e32\u0e1e\u0e31\u0e19\u0e40\u0e01\u0e49\u0e32\u0e23\u0e49\u0e2d\u0e22\u0e40\u0e01\u0e49\u0e32\u0e2a\u0e34\u0e1a\u0e40\u0e01\u0e49\u0e32END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e2b\u0e21\u0e37\u0e48\u0e19END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e40\u0e2d\u0e47\u0e14END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e2a\u0e2d\u0e07END", "BEGIN\u0e2b\u0e19\u0e36\u0e48\u0e07\u0e2b\u0e21\u0e37\u0e48\u0e19\u0e2a\u0e32\u0e21END"],
        "vietnamese-1" => ["BEGINm\u00f4\u0323tEND", "BEGINhaiEND", "BEGINbaEND", "BEGINb\u00f4\u0301nEND", "BEGINn\u0103mEND", "BEGINsa\u0301uEND"],
        "vietnamese-98" => ["BEGINchi\u0301n m\u01b0\u01a1i ta\u0301mEND", "BEGINchi\u0301n m\u01b0\u01a1i chi\u0301nEND", "BEGINm\u00f4\u0323t tr\u0103mEND", "BEGINm\u00f4\u0323t tr\u0103m le\u0309 m\u00f4\u0323tEND", "BEGINm\u00f4\u0323t tr\u0103m le\u0309 haiEND", "BEGINm\u00f4\u0323t tr\u0103m le\u0309 baEND"],
        "vietnamese-998" => ["BEGINchi\u0301n tr\u0103m chi\u0301n m\u01b0\u01a1i ta\u0301mEND", "BEGINchi\u0301n tr\u0103m chi\u0301n m\u01b0\u01a1i chi\u0301nEND", "BEGINm\u00f4\u0323t nga\u0300nEND", "BEGINm\u00f4\u0323tEND", "BEGINhaiEND", "BEGINbaEND"],
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
                { Ascii = "Leelawadee UI", HighAnsi = "Leelawadee UI", EastAsia = "Leelawadee UI" }));
        }
        return stream.ToArray();
    }
}
