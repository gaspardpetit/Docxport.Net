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

public sealed class DocSexagenaryMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "traditional", "zodiac", "cycle" })
            foreach (var pair in new[] { (0, 1), (0, Limit(format + "-1") - 1), (8, Limit(format + "-1") - 1) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void SexagenaryMarkersSurviveDoc(int depth, string mode)
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

    private static (byte Code, NumberFormatValues Format, int Limit) Definition(string mode) =>
        mode[..mode.LastIndexOf('-')] switch
        {
            "traditional" => (0x1E, NumberFormatValues.IdeographTraditional, 10),
            "zodiac" => (0x1F, NumberFormatValues.IdeographZodiac, 12),
            "cycle" => (0x20, NumberFormatValues.IdeographZodiacTraditional, 60),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    private static int Limit(string mode) => Definition(mode).Limit;

    public static string[] Labels(int depth, string mode)
    {
        const string stems = "甲乙丙丁戊己庚辛壬癸";
        // Word uses U+620D for the eleventh branch.
        const string branches = "子丑寅卯辰巳午未申酉戍亥";
        return Enumerable.Range(int.Parse(mode.Split('-').Last()), 6)
            .Select(value => "BEGIN" + (mode.StartsWith("cycle-")
                ? stems[(value - 1) % 10].ToString() + branches[(value - 1) % 12]
                : value > Limit(mode) ? value.ToString()
                : (mode.StartsWith("traditional-") ? stems : branches)[value - 1].ToString()) + "END").ToArray();
    }

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
