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

public sealed class DocIndicMarkerTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var format in new[] { "vowels", "consonants", "numbers", "counting" })
        {
            var alphabetic = format is "vowels" or "consonants";
            var boundaryStart = format == "vowels" ? 36 : format == "consonants" ? 17 : 98;
            var cycleStart = alphabetic ? 303 : 9998;
            foreach (var pair in new[] { (0, 1), (0, boundaryStart), (0, cycleStart), (8, cycleStart) })
                yield return new object[] { pair.Item1, format + "-" + pair.Item2 };
        }
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void IndicMarkersSurviveDoc(int depth, string mode)
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
            "vowels" => (0x31, NumberFormatValues.HindiVowels),
            "consonants" => (0x32, NumberFormatValues.HindiConsonants),
            "numbers" => (0x33, NumberFormatValues.HindiNumbers),
            "counting" => (0x34, NumberFormatValues.HindiCounting),
            _ => throw new ArgumentException("Unknown test format", nameof(mode))
        };

    // Editable ListString values captured from Word source documents on 2026-11-15.
    public static string[] Labels(int depth, string mode) => mode switch
    {
        "consonants-1" => ["BEGIN\u0905END", "BEGIN\u0906END", "BEGIN\u0907END", "BEGIN\u0908END", "BEGIN\u0909END", "BEGIN\u090aEND"],
        "consonants-17" => ["BEGIN\u0905\u0902END", "BEGIN\u0905\u0903END", "BEGIN\u0905\u0905END", "BEGIN\u0906\u0906END", "BEGIN\u0907\u0907END", "BEGIN\u0908\u0908END"],
        "consonants-303" => ["BEGIN\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913\u0913END", "BEGIN\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914\u0914END", "BEGIN\u0905END", "BEGIN\u0906END", "BEGIN\u0907END", "BEGIN\u0908END"],
        "counting-1" => ["BEGIN\u090f\u0915END", "BEGIN\u0926\u094bEND", "BEGIN\u0924\u0940\u0928END", "BEGIN\u091a\u093e\u0930END", "BEGIN\u092a\u093e\u0901\u091aEND", "BEGIN\u091b\u0903END"],
        "counting-98" => ["BEGIN\u0905\u0920\u094d\u0920\u093e\u0928\u0935\u0947END", "BEGIN\u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947END", "BEGIN\u090f\u0915 \u0938\u094cEND", "BEGIN\u090f\u0915 \u0938\u094c \u090f\u0915END", "BEGIN\u090f\u0915 \u0938\u094c \u0926\u094bEND", "BEGIN\u090f\u0915 \u0938\u094c \u0924\u0940\u0928END"],
        "counting-9998" => ["BEGIN\u0928\u094c \u0939\u091c\u093e\u0930 \u0928\u094c \u0938\u094c \u0905\u0920\u094d\u0920\u093e\u0928\u0935\u0947END", "BEGIN\u0928\u094c \u0939\u091c\u093e\u0930 \u0928\u094c \u0938\u094c \u0928\u093f\u0928\u094d\u092f\u093e\u0928\u0935\u0947END", "BEGIN\u090f\u0915END", "BEGIN\u0926\u094bEND", "BEGIN\u0924\u0940\u0928END", "BEGIN\u091a\u093e\u0930END"],
        "numbers-1" => ["BEGIN\u0967END", "BEGIN\u0968END", "BEGIN\u0969END", "BEGIN\u096aEND", "BEGIN\u096bEND", "BEGIN\u096cEND"],
        "numbers-98" => ["BEGIN\u096f\u096eEND", "BEGIN\u096f\u096fEND", "BEGIN\u0967\u0966\u0966END", "BEGIN\u0967\u0966\u0967END", "BEGIN\u0967\u0966\u0968END", "BEGIN\u0967\u0966\u0969END"],
        "numbers-9998" => ["BEGIN\u096f\u096f\u096f\u096eEND", "BEGIN\u096f\u096f\u096f\u096fEND", "BEGIN\u0967\u0966\u0966\u0966\u0966END", "BEGIN\u0967\u0966\u0966\u0966\u0967END", "BEGIN\u0967\u0966\u0966\u0966\u0968END", "BEGIN\u0967\u0966\u0966\u0966\u0969END"],
        "vowels-1" => ["BEGIN\u0915END", "BEGIN\u0916END", "BEGIN\u0917END", "BEGIN\u0918END", "BEGIN\u0919END", "BEGIN\u091aEND"],
        "vowels-303" => ["BEGIN\u091b\u091b\u091b\u091b\u091b\u091b\u091b\u091b\u091bEND", "BEGIN\u091c\u091c\u091c\u091c\u091c\u091c\u091c\u091c\u091cEND", "BEGIN\u0915END", "BEGIN\u0916END", "BEGIN\u0917END", "BEGIN\u0918END"],
        "vowels-36" => ["BEGIN\u0938END", "BEGIN\u0939END", "BEGIN\u0915\u0915END", "BEGIN\u0916\u0916END", "BEGIN\u0917\u0917END", "BEGIN\u0918\u0918END"],
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
                { Ascii = "Nirmala UI", HighAnsi = "Nirmala UI", EastAsia = "Nirmala UI" }));
        }
        return stream.ToArray();
    }
}
